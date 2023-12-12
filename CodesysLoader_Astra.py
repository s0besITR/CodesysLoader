# -*- coding: utf-8 -*-
#############################################################################################################################
#	Вызов скрипта в Astra IDE: Инструменты -> Скрипты -> Выполнить скрипт													#
#																															#
#	После вызова скрипта появится выбор:																					#
#		1) Работа с активным проектом 																						#
#		2) Работа с несколькими проектами																					#
#		3) Импорт modbus устройств в активный проект																		#
#																															#
#	1) При работе с активным проектом следует выбрать один или несколько .xml файлов:										#
#		Application, EDC_CMD, EDC_DATA, ST_CMD, ST_DATA																		#
#	Тогда произойдет импорт выбранных файлов в текущий открытый проект.														#
#																															#
#	2) При выборе второго варианта (несколько проектов) скрипт предложит выбрать общий путь до папок						#
#	в которых лежат .xml файлы. Например выбираем путь																		#
#			C:\Users\Собецкий Александр\Desktop\Codesys UI\out																#
#	Структура внутри:																										#
#		-USO1:																												#
#			- REGUL_Application.xml																							#
#			- REGUL_IEC104_EDC_CMD.iec104cmd.xml																			#
#			- REGUL_IEC104_EDC_DATA.iec104data.xml																			#
#			- REGUL_IEC104_ST_CMD.iec104cmd.xml																				#
#			- REGUL_IEC104_ST_DATA.iec104data.xml																			#
#		-USO2:																												#
#			- REGUL_Application.xml																							#
#		-...																												#
#																															#
#	При работе с несколькими проектами алгоритм следующий:																	#
#	a) Получаем список папок в корне (например USO1, USO2, ... , USON)														#
#	b) Для каждой папки ищем одноименный ПЛК в проекте и если находим, то устанавливаем у него активное приложение			#
#	c) Производим импорт всех найденных xml файлов в папке в активное приложение, как при работе с активным проектом		#
#	d) Переходим к следующей папке по списку и идем на пункт b																#
#																															#
#	В режимах 1 или 2, если рядом с .xml файлами будет папка modules, будет произведен импорт найденных файлов				#
#	DEV_ALL.XML, как при выборе режима 3																					#
#																															#
#	3) Выбор папки modules и импорт найденных файлов DEV_ALL.XML в соответствующие каналы Modbus_Serial_Master				#
#		Перед импортом будут удалены вообще все modbus serial outer slave из проекта активного приложения					#
#		(со всех портов всех устройств)																						#
#		Структура пример:																									#
#		-modules:																											#
#			- Модуль 8:																										#
#				- Канал 1:																									#
#					- DEV_ALL.XML																							#
#				- Канал 2:																									#
#					- DEV_ALL.XML																							#
#			...																												#
#			- Модуль 9:																										#
#			...																												#
#																															#
#	P.S																														#
#	Работа в режиме командной строки, как это было в скрипте CodesysLoader.py для Epsilon LD не предусмотрено				#																										#
#																															#
#	31.05.2022 Собецкий А.В																									#
#############################################################################################################################

import gc, os, warnings, shutil
warnings.filterwarnings("ignore", category=DeprecationWarning)
import xml.etree.ElementTree as ET
from datetime import datetime

start_time = datetime.now()

# Здесь можно переопределить путь по умолчанию в диалоговом окне выбора папки
#main_dir = "C:\\Users\Собецкий Александр\\Desktop\\Терехово Ген\\out"
main_dir = os.environ['USERPROFILE'] + "\\Desktop"

################		ВСПОМОГАТЕЛЬНОЕ			#################
class Reporter(ImportReporter, ExportReporter):
	"""
	Класс для импорта/экспорта объектов
	"""
	def error(self, message):
		write_msg(Severity.Error, message)
	def warning(self, message):
		write_msge(Severity.Warning, message)
	def resolve_conflict(self, obj):
		return ConflictResolve.Replace
	def added(self, obj):
		write_msg(Severity.Text, 'Добавлено:   ' + obj.get_name())
	def replaced(self, obj):
		write_msg(Severity.Text, 'Заменено:   ' + obj.get_name())
	def skipped(self, obj):
		print("Пропущено: ", obj.get_name())
	def nonexportable(self):
		write_msg(Severity.Error, "Файл не экспортируемый")
	@property
	def aborting(self):
		return False

class SilentReporter(ImportReporter, ExportReporter):
	"""
	Класс для импорта/экспорта объектов в безшумном режиме (без вывода в лог)
	"""
	def error(self, message):
		write_msg(Severity.Error, message)
	def warning(self, message):
		write_msge(Severity.Warning, message)
	def resolve_conflict(self, obj):
		return True
	def added(self, obj):
		return True
	def replaced(self, obj):
		return True
	def skipped(self, obj):
		return True
	def nonexportable(self):
		write_msg(Severity.Error, "Файл не экспортируемый")
	@property
	def aborting(self):
		return False

def getPrjPath():
	"""
		Возвращает путь до каталога открытого файла проекта
	"""
	tmp = projects.primary.path.split('\\')
	tmp.pop(-1)
	return '\\'.join(tmp)			

def getActiveApplication():
	"""
	Возвращает объект активного приложения в проекте
	"""	
	return projects.primary.active_application
	
def getActiveDevice():
	"""
	Получить объект устройства в активном проекте
	"""
	return getActiveApplication().parent.parent

def getActiveCrate():
	"""
	Возвращаем объект крейта. Сначала ищем модуль A2, затем берем его родителя
	"""
	obj = getActiveDevice().find('A2', True)
	if len(obj) > 0 :
		return obj[0].parent
	return None

def write_msg(severity, msg):
	"""
	Записывает сообщение с указанным приоритетом на стандартный вывод консоли Codesys.
	А так же в лог файл на всякий случай
	"""
	system.write_message(severity, msg)
	with open(getPrjPath() + "\\CodesysLoader_log.txt", 'a') as log:
		if severity == Severity.Error:
			msg = 'ERROR:\t' + msg
		if severity == Severity.Warning:
			msg = 'Warn:\t' + msg
		if severity == Severity.Information:
			msg = 'Info:\t' + msg
		if severity == Severity.Text:
			msg = '\t' + msg
		msg += '\n'
		log.write(msg.encode('utf8'))

def set_primary_prj(prim_name):
	"""
	Установить активный проект на имя prim_name
	"""
	for prj in projects.all:
		for obj in prj.get_children():
			if(obj.is_device):				
				if obj.get_name(False) == prim_name :					
					app = obj.find('Application', recursive = True)
					if len(app) > 0 :
						prj.active_application = app[0]
						write_msg(Severity.Information, "Сменили активное приложение на {0}".format(prim_name))
						return True
					else:
						write_msg(Severity.Warning, "Устройство {0} найдено, но внутри нет объекта Application".format(prim_name))
						return False
	write_msg(Severity.Warning,"Не удалось поменять активное приложение. {0} не найдено".format(prim_name))
	return False

def compileActivePrj():
	"""
	Коммилирует активный проект, возвращает кол-во ошибок, возникших при компиляции
	"""
	app = getActiveApplication()
	app.rebuild()
	guid_compile_category = Guid("{97F48D64-A2A3-4856-B640-75C046E37EA9}")
	messages = system.get_message_objects(guid_compile_category)
	errors = 0
	for msg in messages:
		if msg.severity == Severity.FatalError or msg.severity == Severity.Error:
			errors += 1
	if errors != 0:
		write_msg(Severity.Error, 'Компиляция завершилась с ошибками. Ошибок: ' + str(errors))
	else:
		write_msg(Severity.Text, 'Компиляция завершилась успешно')
	return errors	

################		ИМПОРТ В ПРОЕКТ			#################

def modifyActiveProject(files):
	"""
	Обновляет активный проект. Загружает Application, обновляет Slave104Drivers, ModbusSlaveDriver
	"""
	#Списки каналов
	global start_time
	
	slave_ST = {}
	slave_EDC = {}
	slave_Modbus = dict.fromkeys(['data']);
	
	if start_time == 0:
		start_time = datetime.now()
			
	
	if files != None:
		for file in files:
			if file.endswith('Application.xml'):
				import_application(file)
			elif 'REGUL_IEC104_' in file:
				key_str = ''
				filename = file[file.rfind('\\') + 1:]
				if filename.endswith('iec104cmd.xml'):
					key_str = 'cmd'
				if filename.endswith('iec104data.xml'):
					key_str = 'data'				
				if '_EDC_' in filename:
					slave_EDC[key_str + filename.replace('REGUL_IEC104_', '').replace('_CMD.iec104cmd.xml','').replace('_DATA.iec104data.xml','').replace('ST','').replace('EDC','')] = load_channels(file)
				if '_ST_' in filename:
					slave_ST[key_str + filename.replace('REGUL_IEC104_', '').replace('_CMD.iec104cmd.xml','').replace('_DATA.iec104data.xml','').replace('ST','').replace('EDC','')] = load_channels(file)
			elif file.endswith('mb_direct_channels.xml'):
				slave_Modbus['data'] = load_channels_modbus(file) 
			elif not (file.endswith('modules') or file.endswith('modules_tcp') ):				
				write_msg(Severity.Error, 'Неизвестный формат файла:   ' + file)
	else:
		write_msg(Severity.Text, 'Выполнение скрипта отменено')
		return False
	
	if slave_ST:
		clearOld_GVL('I104_GVL_TM')
		for i in range(1,150):
			index = str(i)
			new_dict = {}
			module_name = ''
			gvl_name = ''
			if index == '1':
				for key in ['data', 'cmd']:
					if key in slave_ST:
						new_dict[key] = slave_ST[key]
				
				module_name = 'Slave_104_Driver'
				gvl_name = 'I104_GVL_TM'
			else:
				for key in ['data_' + index, 'cmd_' + index]:
					if key in slave_ST:
						new_dict[key.replace('_' + index,'')] = slave_ST[key]
				module_name = 'Slave_104_Driver_' + index
				gvl_name = 'I104_GVL_TM_' + index
			if new_dict:
				iec104slave_mod(module_name, new_dict)				
				iec104_GVL(gvl_name, new_dict)	
			else:
				break
		write_msg(Severity.Text, '----------Импорт каналов ТМ закончен----------')
		
	if slave_EDC:
		clearOld_GVL('I104_GVL_KK')
		for i in range(1,150):
			index = str(i)
			new_dict = {}
			module_name = ''
			gvl_name = ''
			if index == '1':
				for key in ['data', 'cmd']:
					if key in slave_EDC:
						new_dict[key] = slave_EDC[key]
				
				module_name = 'Slave_104_Driver_EDC'
				gvl_name = 'I104_GVL_KK'
			else:
				for key in ['data_' + index, 'cmd_' + index]:
					if key in slave_EDC:
						new_dict[key.replace('_' + index,'')] = slave_EDC[key]
				module_name = 'Slave_104_Driver_EDC_' + index
				gvl_name = 'I104_GVL_KK_' + index
			if new_dict:
				iec104slave_mod(module_name, new_dict)				
				iec104_GVL(gvl_name, new_dict)	
			else:
				break
		write_msg(Severity.Text, '----------Импорт каналов EDC закончен----------')
		
		
	if slave_Modbus.get('data')!= None:
		modbus_slave_mod('Modbus_Tcp_Slave', slave_Modbus)		
		write_msg(Severity.Text, '----------Импорт каналов ModbusTCPSlave закончен----------')
		
	dir = files[0].split('\\')
	dir.pop(-1)
	replaceModbusDevices('\\'.join(dir))
	commentTriggers()
	# очистка памяти
	#del slave_ST
	#del slave_EDC
	del slave_Modbus
	gc.collect()
	return True	

def modifyManyProjects(folder):
	"""
	Обновляет все проекты скопом по очереди
	"""
	err_count = 0	
	for uso_folder in os.listdir(folder):	
		dir = folder + '\\' + uso_folder
		xml_list = []										# список .xml для импорта
		if os.path.isdir(dir):
			file_list =  os.listdir(dir)			
			for file in file_list:
				if (file.endswith('.xml') or file == "modules" or file == "modules_tcp"):
					xml_list.append(dir + '\\' + file)				
		if len(xml_list):
			if(set_primary_prj(uso_folder)):			
				modifyActiveProject(xml_list)			
				err_count += compileActivePrj()
				write_msg(Severity.Information, 'Работа с текущим устройством завершена')			
	return err_count

### Импорт XML	
def import_xml(xml_name):
	"""
		Импорт XML файла в проект
	"""
	app = getActiveApplication()
	
	if not app is None :
		app.import_xml(Reporter(), xml_name, True)	
	
	### Импорт Application	
def import_application(xml_name):
	"""
		Импорт файла Application в проект
	"""
	app = getActiveApplication()
	
	if not app is None :
		app.import_xml(Reporter(), xml_name)
		prettify_imitation_prg()
		write_msg(Severity.Text, '----------Импорт Application закончен----------')
	else:
		write_msg(Severity.Error, '----------Импорт application не состоялся----------')

### Модификация Драйвера Modbus TCP Slave
def load_channels_modbus(xml_name):
	""" 
		Парсим xml файл, и возвращает список словарей, в которых хранятся атрибуты канала
	"""	
	# Складируем информацию из xml в список объектов
	tmp_list = []	
	tree = ET.parse(xml_name)
	root = tree.getroot()
	for elem in root:		
		tmp_list.append(
					{
						"Name" 			: elem.attrib.get('Name'),
						"Descr" 		: elem.attrib.get('Descr'),
						"Type" 			: elem.attrib.get('Type'),
						"Offset" 		: elem.attrib.get('Offset'),
						"Length" 		: elem.attrib.get('Length'),
						"VarName" 		: elem.attrib.get('VarName')
					}		
				)				
	return tmp_list
	
def modbus_slave_mod(driver_name, channels):
	"""
	Экспортируем из Codesys Modbus Tcp Slave с именем driver_name, затем копируем его построчно в новый файл.
	Когда нашли каналы в старом файле, пропускаем их, вместо них копируем новые каналы, сохраняем файл с припиской _mod
	"""
	# Экспорт файла
	obj = getActiveDevice().find(driver_name, True)
	file_name = getPrjPath() + "\\" + driver_name + ".xml"
	file_name_mod = file_name.replace('.xml', '_mod.xml')
	if len(obj) > 0 :
		obj[0].export_xml(Reporter(), file_name)		
		obj[0].remove()
		write_msg(Severity.Text, 'Удалено:   {driver}'.format(driver = driver_name))
	else:
		write_msg(Severity.Error, '{driver} не найден!'.format(driver = driver_name))
		return 
	
	# Модификация файла
	channels_found = False
	with open(file_name, 'rb') as src, open(file_name_mod, 'wb') as dest:
		for line in src:
			if line.find('localTypes:Channel') != -1 and channels_found == False:
				channels_found = True
			if line.find('</HostParameterSet>')!= -1:
				writeModbusChannelsToFile(dest, channels)
				channels_found = False
			if channels_found == False:
				dest.write(line)
		
	dev = getActiveDevice()
	if (dev != None):
		dev.import_xml(Reporter(), file_name_mod)
		os.remove(file_name)
		os.remove(file_name_mod)
	else:
		write_msg(Severity.Error, 'Устройство Device не найдено!')	

def writeModbusChannelsToFile(file, channels):
	"""
	Записывает в файл Modbus TCP Slave все ранее сохраненные каналы
	"""	
	param_id = 20000								# id данных начинается с 20000	
	if(channels['data']):
		for ch in channels['data']:	
			file.write(getModbusNodeStr(ch, str(param_id)))
			param_id += 1	

def getModbusType(type_str):
	"""
	Получить тип канала в виде числа
	"""	
	type_ids = {		
		'4' :'HoldingRegisters',
		'1' :'DiscreteInputs',
		'3' : 'InputRegisters',
		'2' : 'Coils'		
	}
	
	for key in type_ids.keys():
		if type_str == type_ids[key]:
			return key
	return 'None'
	
def getModbusNodeStr(channel, param_id):
	"""
	Возвращаем xml представление одного modbus канала с данными
	"""	
	f_str = """<Parameter ParameterId="{paramId}" type="localTypes:Channel">
		<Attributes />
		<Value name="{valName}" visiblename="{valVisName}" desc="{valDesc}">
			<Element name="ChType" visiblename="Тип канала">{valType}</Element>
			<Element name="Offset" visiblename="Смещение">{valOffset}</Element>
			<Element name="Length" visiblename="Длина">{valLength}</Element>
			<Element name="VarName" visiblename="Имя переменной">'{var_name}'</Element>
			<Element name="ChannelName" visiblename="Имя">'{valVisName}'</Element>
			<Element name="ChannelComment" visiblename="Комментарий">'{valDesc}'</Element>
		</Value>
		<Name>{valVisName}</Name>
		<Description>{valDesc}</Description>
	</Parameter>"""
	
	
	node_str = f_str.format(
		paramId 		= param_id,
		valName 		= '_x003' + param_id[0] + '_' + param_id[1:],
		valVisName 		= channel.get('Name'),
		valDesc 		= channel.get('Descr'),
		valType 		= getModbusType(channel.get('Type')),
		valOffset		= channel.get('Offset'),
		valLength		= channel.get('Length'),
		var_name 		= channel.get('VarName')		
		)
	return node_str.encode('utf-8')
		
	### Модификация Драйвера Slave 104		
def load_channels(xml_name):
	""" 
		Парсим xml файл, и возвращает список словарей, в которых хранятся атрибуты канала
	"""	
	# Складируем информацию из xml в список объектов
	tmp_list = []	
	tree = ET.parse(xml_name)
	root = tree.getroot()
	for elem in root:		
		tmp_list.append(
					{
						"Name" 			: elem.attrib.get('Name'),
						"Descr" 		: elem.attrib.get('Descr'),
						"TypeId" 		: elem.attrib.get('TypeId'),
						"CustomTypeId" 	: elem.attrib.get('CustomTypeId'),
						"AutoTime" 		: elem.attrib.get('AutoTime'),
						"HighBound" 	: elem.attrib.get('HighBound'),
						"LowBound" 		: elem.attrib.get('LowBound'),
						"Scale" 		: elem.attrib.get('Scale'),
						"IoAdr" 		: elem.attrib.get('IoAdr'),
						"MapVarName" 	: elem.attrib.get('MapVarName'),
						"MirrorAdr" 	: elem.attrib.get('MirrorAdr'),
						"SelectPeriod" 	: elem.attrib.get('SelectPeriod'),
						"ExecTimeout" 	: elem.attrib.get('ExecTimeout'),			
						"Cycle" 		: elem.attrib.get('Cycle'),
						"DeadBand" 		: elem.attrib.get('DeadBand'),
						'lib_type'		: getLibType(
													int(elem.attrib.get('TypeId')),
													int(elem.attrib.get('CustomTypeId'))
										)
					}		
				)				
	return tmp_list
	
def getLibType(type_id, c_type_id):
	"""
	По идентификатору параметра получаем библиотечный тип
	"""
	# Принадлежность библиотечному типу к идентификатору из TypeId
	type_ids = {		
		'c_bo_na_1_fb' : [51],
		'c_bo_ta_1_fb' : [64],
		'c_dc_na_1_fb' : [46],
		'c_dc_ta_1_fb' : [59],
		'c_rc_na_1_fb' : [47],
		'c_rc_ta_1_fb' : [60],
		'c_sc_na_1_fb' : [45],
		'c_sc_ta_1_fb' : [58],
		'c_se_na_1_fb' : [48],
		'c_se_nb_1_fb' : [49],
		'c_se_nc_1_fb' : [50],
		'c_se_ta_1_fb' : [61],
		'c_se_tb_1_fb' : [62],
		'c_se_tc_1_fb' : [63],		
	
		'm_bo_tb_1_fb' : [3, 5, 7, 31, 32, 33],
		'm_ep_td_1_fb' : [38],
		'm_it_tb_1_fb' : [15, 37],
		'm_me_tf_1_fb' : [9, 11, 13, 21, 34, 35, 36],
		'm_sp_tb_1_fb' : [1, 30]
	}
	# Принадлежность библиотечному типу к идентификатору из CustomTypeId
	custom_type_ids = {
		'Lreal_tc_fb' : [1, 2],
		'ulint_tc_fb' : [3, 4]
	}	
	for key in custom_type_ids.keys():
		if c_type_id in custom_type_ids[key]:
			return key
	for key in type_ids.keys():
		if type_id in type_ids[key]:
			return key
	return 'None'
	
def iec104slave_mod(driver_name, channels):
	"""
	Экспортируем из Codesys Slave104Driver с именем driver_name, затем копируем его построчно в новый файл.
	Когда нашли каналы в старом файле, пропускаем их, вместо них копируем новые каналы, сохраняем файл с припиской _mod
	"""
	# Экспорт файла
	obj = getActiveDevice().find(driver_name, True)
	file_name = getPrjPath() + "\\" + driver_name + ".xml"
	file_name_mod = file_name.replace('.xml', '_mod.xml')
	if len(obj) > 0 :
		obj[0].export_xml(Reporter(), file_name)		
		obj[0].remove()
		write_msg(Severity.Text, 'Удалено:   {driver}'.format(driver = driver_name))
	else:
		write_msg(Severity.Error, '{driver} не найден!'.format(driver = driver_name))
		return 
	
	# Модификация файла
	channels_found = False
	with open(file_name, 'rb') as src, open(file_name_mod, 'wb') as dest:
		for line in src:
			if line.find('localTypes:iec101') != -1 and channels_found == False:
				channels_found = True
			if line.find('</HostParameterSet>')!= -1:
				writeChannelsToFile(dest, channels)
				channels_found = False
			if channels_found == False:
				dest.write(line)
		
	dev = getActiveDevice()
	if (dev != None):
		dev.import_xml(Reporter(), file_name_mod)
		os.remove(file_name)
		os.remove(file_name_mod)
	else:
		write_msg(Severity.Error, 'Устройство Device не найдено!')

def writeChannelsToFile(file, channels):
	"""
	Записывает в файл Slave104Driver все ранее сохраненные каналы
	"""	
	param_id = 64000								# id данных начинается с 64000	
	if(channels['data']):
		for ch in channels['data']:	
			file.write(getDataNodeStr(ch, str(param_id)))
			param_id += 1
	param_id = 74000								# id команд с 74000
	if(channels['cmd']):
		for ch in channels['cmd']:	
			file.write(getCmdNodeStr(ch, str(param_id)))
			param_id += 1

def getDataNodeStr(channel, param_id):
	"""
	Возвращаем xml представление одного канала с данными
	"""	
	f_str = """<Parameter ParameterId="{paramId}" type="{paramType}">
		<Attributes onlineaccess="read" />
		<Value name="{valName}" visiblename="{valVisName}" onlineaccess="read" desc="{valDesc}">
			<Element name="typeid" visiblename="Type ID">{typeid}</Element>
			<Element name="io_addr" visiblename="Information object address">{io_addr}</Element>
			<Element name="cycle" visiblename="Subject for cyclic transmission">{cycle}</Element>
			<Element name="deadband" visiblename="Dead band">{deadband}</Element>
			<Element name="hi" visiblename="High bound">{hi}</Element>
			<Element name="lo" visiblename="Low bound">{lo}</Element>
			<Element name="coef" visiblename="Scale coefficient">{coef}</Element>
			<Element name="var_name" visiblename="Variable name">'{var_name}'</Element>
			<Element name="custom_typeid" visiblename="Custom typeID">{custom_typeid}</Element>
		</Value>
		<Name>{Name}</Name>
		<Description>{Description}</Description>
    </Parameter>"""	
	node_str = f_str.format(
		paramId 		= param_id,
		paramType 		= 'localTypes:iec101data_new_descr',		
		valName 		= '_x003' + param_id[0] + '_' + param_id[1:],
		valVisName 		= channel.get('Name'),
		valDesc 		= channel.get('Descr'),		
		typeid 			= channel.get('TypeId'),
		io_addr 		= channel.get('IoAdr'),
		cycle 			= channel.get('Cycle'),
		deadband 		= addFloatNumbers(channel.get('DeadBand')),
		hi 				= addFloatNumbers(channel.get('HighBound')),
		lo 				= addFloatNumbers(channel.get('LowBound')),
		coef 			= addFloatNumbers(channel.get('Scale')),
		var_name 		= channel.get('MapVarName'),
		custom_typeid 	= channel.get('CustomTypeId'),		
		Name			= channel.get('Name'),
		Description 	= channel.get('Descr')
		)
	return node_str

def getCmdNodeStr(channel, param_id):
	"""
	Возвращаем xml представление одного канала с командами
	"""	
	f_str = """<Parameter ParameterId="{paramId}" type="{paramType}">
		   <Attributes onlineaccess="read" />
		   <Value name="{valName}" visiblename="{valVisName}" onlineaccess="read" desc="{valDesc}">
				<Element name="typeid" visiblename="Type ID">{typeid}</Element>
				<Element name="io_addr" visiblename="Information object address">{io_addr}</Element>
				<Element name="mirror_addr" visiblename="Address for send back mirror var value">{mir_adr}</Element>
				<Element name="scale" visiblename="Scale">{scale}</Element>
				<Element name="lo" visiblename="Low border">{lo}</Element>
				<Element name="hi" visiblename="Hi border">{hi}</Element>
			   <Element name="select_period" visiblename="Select period in seconds">{per}</Element>
			   <Element name="exec_timeout" visiblename="Execution timeout in seconds">{timeout}</Element>
			   <Element name="var_name" visiblename="Variable name">'{var_name}'</Element>
		   </Value>
		   <Name>{Name}</Name>
		   <Description>{Description}</Description>
       </Parameter>
	   """	   
	node_str = f_str.format(
		paramId 		= param_id,
		paramType 		= 'localTypes:iec101cmd_new_descr',		
		valName 		= '_x003' + param_id[0] + '_' + param_id[1:],
		valVisName 		= channel.get('Name'),
		valDesc 		= channel.get('Descr'),		
		typeid 			= channel.get('TypeId'),
		io_addr 		= channel.get('IoAdr'),
		mir_adr 		= channel.get('MirrorAdr'),
		scale 			= addFloatNumbers(channel.get('Scale')),
		lo 				= addFloatNumbers(channel.get('LowBound')),
		hi 				= addFloatNumbers(channel.get('HighBound')),
		per				= channel.get('SelectPeriod'),
		timeout			= channel.get('ExecTimeout'),
		var_name 		= channel.get('MapVarName'),
		Name			= channel.get('Name'),
		Description 	= channel.get('Descr')
		)
	return node_str

def addFloatNumbers(param):
	"""
	Добивает числовой параметр до 10ти знаков после запятой
	"""
	delta = 10 - (len(param) - len(param[0:param.find(".")]) - 1)
	result = (param + '.' + '0'*delta) if delta == 10 else param + '0'*delta
	return result
	
	### Создание GVL по каналам в драйвере Slave 104	
def iec104_GVL(gvl_name, channels):
	""""
	Добавляем новый gvl. Наполняем переменными из каналов
	"""	
		
	# Добавляем пустой новый
	app = getActiveApplication()
	if not app is None :
		obj = app.find('GVL_I104')
		if len(obj) == 0 :
			app.create_folder('GVL_I104')
		obj = app.find('GVL_I104')	

		new_gvl_obj = obj[0].create_gvl(gvl_name)
		new_gvl_obj.textual_declaration.remove(0, 0, new_gvl_obj.textual_declaration.length)
		new_gvl_obj.textual_declaration.insert(0, 0, getGVLdata(channels))
		write_msg(Severity.Text, 'Добавлено:   {gvl}'.format(gvl = gvl_name))
	else:
		write_msg(Severity.Error, 'Application не найдено')
		return	

def clearOld_GVL(gvl_name):
	"""
	Удаляем все старые GVL
	"""
	
	write_msg(Severity.Text, 'Удаляем все старые GVL:')
	
	gvl_list = [gvl_name] 
	for i in range(2,150):
		gvl_list.append(gvl_list[0] + '_' + str(i))
		
	app = getActiveApplication()
	if not app is None :
		for g in gvl_list:
			obj = app.find(g, True)			
			if len(obj) > 0 :
				obj[0].remove()
				write_msg(Severity.Text, 'Удалено:   {gvl}'.format(gvl = g))
		
def getGVLdata(channels):
	"""
	Из каналов получаем внутренний текст GVL
	"""		
	gvl_text = "VAR_GLOBAL" + '\n' + '\t//Updated: ' + datetime.now().strftime('%d.%m.%Y %H:%M:%S') + '\n'
	if(channels['data']):
		for ch in channels['data']:		
			gvl_text += '\t// ' + ch.get('Descr')[1:] + '\n\t' + ch.get('MapVarName') + ': PsIEC60870Bridge.' + ch.get('lib_type') + ';\n'
	if(channels['cmd']):	
		for ch in channels['cmd']:
			gvl_text += '\t// ' + ch.get('Descr')[1:] + '\n\t' + ch.get('MapVarName') + ': PsIEC60870Bridge.' + ch.get('lib_type') + ';\n'
	gvl_text+= 'END_VAR'
	return gvl_text

	### Секция IMIT. Преобразуем ее так, что бы было более функционально
def prettify_imitation_prg():
	"""
		USOX_IMITATION_Imit файл делаем вид более понятным за счет группировки строчек по объектам
	"""
	imit = getImitationObj()
	imit_text = ""
	initFlag = False
	obj_dict = {}
	obj_keys = []
	if imit != None:
		for line_index in range(0,imit.textual_implementation.linecount):
			line_text = imit.textual_implementation.get_line(line_index)			
			if line_text.startswith('IF dInitFlag <>'):
				imit_text += '\n' + line_text
				initFlag = True
			elif line_text.startswith('	dInitFlag :='):
				for k in obj_keys:
					imit_text += "\n\t(* " + k.replace('imit_','') + " *)" + '\n'
					imit_text += '\n'.join(obj_dict[k]) + '\n'
				imit_text += '\n' + line_text
				initFlag = False				
			elif initFlag == True:
				obj_key = getLineImitObj(line_text)			
				if not obj_key in obj_keys:
					obj_keys.append(obj_key)
					obj_dict[obj_key] = []	
				if not " резерв" in line_text.lower(): 			
					obj_dict[obj_key].append(line_text)
			else:
				imit_text += '\n' + line_text
					
		imit.textual_implementation.remove(0, 0, imit.textual_implementation.length)
		imit.textual_implementation.insert(0, 0, imit_text)
		write_msg(Severity.Text, 'Отформатирован:   ' + imit.get_name(False))

def getImitationObj():
	"""
		Поиск файла имитации в проекте
	"""
	app = getActiveApplication()
	if not app is None :	
		for obj in app.get_children():
			if obj.get_name(False).endswith('_IMITATION_Imit'):			
				return obj
	return None

def getLineImitObj(line):
	"""
		Вытащить у строки текст из скобок [текст]
	"""
	return line[line.find('[')+1:line.find(']')]

def commentTriggers():
	"""
	Составляя очередь мы складировали триггеры в GVL файлах, теперь нужно удалить дубликаты из gvl файла triggers, который импортируется вместе с Application
	"""
	obj = getActiveDevice().find("triggers", True)
	if len(obj) == 0 :
		return
		
	#Проходимся по всем добавленным GVL портов и складируем все триггеры в списке port_triggers
	obj = getActiveDevice().find("ModbusControlGVLs", True)
	line_text = ""
	if len(obj) == 0 :
		return
	
	write_msg(Severity.Text, '----------Найдена папка ModbusControlGVLs, проверяем triggers на пересечения----------')
		
	port_triggers = []
	for gvl_file in obj[0].get_children(False):
		for line_index in range(0,gvl_file.textual_declaration.linecount):
			line_text = gvl_file.textual_declaration.get_line(line_index)
			if ("_trigger_" in line_text) and (not "_trigger_EMPTY_" in line_text):
				port_triggers.append(line_text[line_text.index('_'):line_text.index(':')])
					
	#Находим GVL с именем triggers и комментируем те тригера, которые уже есть (из списка port_triggers)		
	obj = getActiveDevice().find("triggers", True)
		
	port_triggers = [x.lower() for x in port_triggers]
	content = ""
	replace_count = 0
	delete_triggers = True
	for line_index in range(0,obj[0].textual_declaration.linecount):
		line_text = obj[0].textual_declaration.get_line(line_index)
		if ("_trigger_" in line_text) and (not "//" in line_text):
			trig_name = line_text[line_text.index('_'):line_text.index(':')]				
			if trig_name.lower() in port_triggers:
				line_text = line_text.replace(trig_name, "//" + trig_name)
				replace_count = replace_count + 1
			else:
				delete_triggers = False
		content = content + line_text + "\n"
	
	if delete_triggers == True:
		obj[0].remove()
		write_msg(Severity.Text, '----------Удален объект triggers, как не содержащий уникальных элементов ----------')
	elif replace_count > 0:
		obj[0].textual_declaration.replace(content)
		write_msg(Severity.Text, '----------Закомментили дублирующие триггеры в triggers. Объект не удален, так как есть уникальные элементы----------')

def import_queue(path):
	"""
	Импортируем файлы из папки Queue
	"""
	obj = getActiveDevice().find("ModbusControlGVLs", True)
	if len(obj) > 0 :
		obj[0].remove()
		write_msg(Severity.Text, 'Удалена папка:   ModbusControlGVLs')
	if os.path.exists(path + '\modules\Queue'):
		write_msg(Severity.Text,'Найдена папка modules\Queue. Производим импорт конфигурации очереди Modbus')		
		for xml in os.listdir(path + '\modules\Queue'):
			import_xml(path + '\\modules\\Queue\\' + xml)		
		write_msg(Severity.Text, '----------Импорт файлов очереди закончен----------')		

	### Импорт модбас устройств
def replaceModbusDevices(path):
	"""
	По выбранному пути path ищет папку modules и импортирует во всех найденные модулях в каждый порт свой DEV_ALL.XML
	"""
	if os.path.exists(path + '\modules'):
		write_msg(Severity.Text,'Найдена папка modules. Перед импортом удаляем все имеющиеся modbus устройства:')
		clearAllModbusDevices()
		for module_folder in os.listdir(path + '\modules'):
			cur_dir = path + '\modules' + '\\' + module_folder
			module = module_folder.replace('Модуль ', 'A')
			obj = getActiveDevice().find(module, True)
			if module_folder == "Queue":
				continue
			if len(obj) > 0 :
				ports = {(child.index + 1) : child for child in obj[0].get_children(False)}
				for ch in os.listdir(cur_dir):					
					cur_dir = path + '\modules' + '\\' + module_folder + '\\' + ch
					ch_object = ports.get(int(ch.replace('Канал ','')))
					if ch_object == None:
						write_msg(Severity.Warning, 'Найден канал ' + ch + ', который отсутствует в модуле ' + module)
					else:
						if 'DEV_ALL.XML' in os.listdir(cur_dir):
							added_devs = importModbusDevices(ch_object, cur_dir + '\\DEV_ALL.XML')
							if len(added_devs) > 0 :
								write_msg(Severity.Text, 'Модуль ' + ch_object.parent.get_name() + ' порт ' + str(ch_object.index + 1) + '.' + ' Добавлено: ' + ', '.join(added_devs))								
						else:
							write_msg(Severity.Warning, 'В канале ' + ch + ' модуля ' + module + ' не найден DEV_ALL.XML. Импорт выполнятся не будет')
			else:
				write_msg(Severity.Warning, 'Найдена папка ' + module_folder + ', но соответствующий модуль не найден в проекте')
		write_msg(Severity.Text, '----------Импорт Modbus устройств закончен----------')
		import_queue(path)				
		return True
	elif os.path.exists(path + '\modules_tcp'):
		write_msg(Severity.Text,'Найдена папка modules_tcp. Перед импортом удаляем все имеющиеся modbus устройства:')
		clearAllModbusTCPDevices()
		obj = getActiveDevice().find('Modbus_TCP_Master', True)
		if obj == None:
			write_msg(Severity.Warning, 'При попытке добавить modbus устройства не был найден модуль Modbus_TCP_Master.')						
		cur_dir = path + '\modules_tcp'
		if 'DEV_ALL.XML' in os.listdir(cur_dir):
			added_devs_tcp = importModbusTcpDevices(obj[0], cur_dir + '\\DEV_ALL.XML')
			if len(added_devs_tcp) > 0 :
				write_msg(Severity.Text, 'Модуль Modbus_TCP_Master. Добавлено: ' + ', '.join(added_devs_tcp))			
		else:
			write_msg(Severity.Warning, 'Не найден DEV_ALL.XML. Импорт выполнятся не будет')
		write_msg(Severity.Text, '----------Импорт ModbusTCP устройств закончен----------')
		return True		
	else:
		write_msg(Severity.Text, 'Этап замены Modbus модулей пропущен, не найдена папка modules')
		return False

def clearAllModbusTCPDevices():
	"""
	Очищает Modbus_TCP_Master модуль в проекте от дочерних устройств.
	"""
	deleted_devs = []
	obj = getActiveDevice().find('Modbus_TCP_Master', True)
	if (obj != None) > 0 :
		for modbus_dev in obj[0].get_children(False):
			deleted_devs.append(modbus_dev.get_name())
			modbus_dev.remove()
		if len(deleted_devs)>0:
			write_msg(Severity.Text,'Модуль Modbus_TCP_Master. Удалено: ' + ', '.join(deleted_devs))	
		write_msg(Severity.Text, '----------Удаление Modbus устройств завершено----------')
	else:
		write_msg(Severity.Warning, 'При попытке очистить modbus устройства не был найден модуль Modbus_TCP_Master.')	
		
def clearAllModbusDevices():
	"""
	Очищает все Modbus Serial Master в проекте от дочерних устройств.
	"""
	crate = getActiveCrate()
	if (crate != None) > 0 :
		for A_module in crate.get_children(False):
			for serial_port in A_module.get_children(False):
				for modbus_port in serial_port.get_children(False):
					if modbus_port.get_name().startswith('Modbus_Serial_Master'):
						deleted_devs = clearModbusDevices(modbus_port)
						if len(deleted_devs)>0:
							write_msg(Severity.Text,'Модуль ' + A_module.get_name() + ' порт ' + str(serial_port.index + 1) + '. Удалено: ' + ', '.join(deleted_devs))
		write_msg(Severity.Text, '----------Удаление Modbus устройств завершено----------')
	else:
		write_msg(Severity.Warning, 'При попытке очистить modbus устройства не был найден крейт.')
	
def clearModbusDevices(modbus_port):
	"""
	Очищает указанный Modbus Serial Master от дочерних устройств. Возвращает список имен удаленных объектов
	"""
	cleared_devs = []
	for dev in modbus_port.get_children(False):
		cleared_devs.append(dev.get_name())
		dev.remove()	
	return cleared_devs

def importModbusTcpDevices(mod, file):
	"""
	Производит импорт в указанный Модуль Modbus_TCP_Master xml файла DEV_ALL.XML. Возвращает список имен добавленных объектов
	"""
	added_devs = []
	if mod == None:
		return added_devs
		
	mod.import_xml(SilentReporter(), file)
	added_devs = [dev.get_name() for dev in mod.get_children()]
	#added_devs.reverse()	
	return added_devs
	
def importModbusDevices(port, file):
	"""
	Производит импорт в указанный Modbus Serial Master xml файла DEV_ALL.XML. Возвращает список имен добавленных объектов
	"""
	children = port.get_children()
	added_devs = []
	if len(children) > 0:
		children[0].import_xml(SilentReporter(), file)
		added_devs = [ch.get_name() for ch in children[0].get_children()]
	else:
		write_msg(Severity.Warning, 'Не найдено устройство Modbus Serial Master для импорта в ' + port.parent.get_name() + ' порт ' + str(port.index + 1))
	added_devs.reverse()
	return added_devs		

################		НАЧАЛЬНОЕ ДИАЛОГОВОЕ ОКНО			#################	
if __name__ == '__main__':
	"""
	Главный UI
	"""
	log_path = getPrjPath() + "\\CodesysLoader_log.txt"
	if os.path.exists(log_path):
		os.remove(log_path)
	res = system.ui.choose("Выберите действие:", ("Работа с активным проектом", "Работа с несколькими проектами", "Импорт modbus устройств в активный проект"))
	if (res[0] == 0):
		if not projects.primary:
			system.ui.error("Не открыто ни одного проекта")
		else:		
			files = system.ui.open_file_dialog("Выберите XML файлы для импорта", directory = main_dir, filter="(*.xml|*.xml", multiselect = True)
			if modifyActiveProject(files):
				write_msg(Severity.Information, 'Завершено!')
				write_msg(Severity.Text, 'Время выполнения: {}'.format(datetime.now() - start_time))
				system.ui.info("Выполнено успешно!")
	elif (res[0] == 1):
		folder = system.ui.browse_directory_dialog("Выберите корневое расположение папок USO1 ... USOx", path = main_dir)
		if folder != None:			
			error_count = modifyManyProjects(folder)
			write_msg(Severity.Information, 'Завершено!')
			write_msg(Severity.Text, 'Время выполнения: {}'.format(datetime.now() - start_time))
			if error_count == 0:
				write_msg(Severity.Information, "Выполнено успешно!")
				system.ui.info("Выполнено успешно!")				
			else:
				write_msg(Severity.Error, "Выполнено с ошибками. Общих ошибок: " + str(error_count))
				system.ui.error("Выполнено с ошибками. Общих ошибок: " + str(error_count))
		else:
			write_msg(Severity.Text, 'Выполнение скрипта отменено')
	elif (res[0] == 2):
		if not projects.primary:
			system.ui.error("Не открыто ни одного проекта")
		else:
			folder = system.ui.browse_directory_dialog("Выберите папку modules (или modules_tcp)", path = main_dir)
			if folder != None:
				if replaceModbusDevices(folder.replace('\\modules_tcp', '').replace('\\modules', '')):
					commentTriggers()
					write_msg(Severity.Information, 'Завершено!')
					write_msg(Severity.Text, 'Время выполнения: {}'.format(datetime.now() - start_time))
					system.ui.info("Выполнено успешно!")
			else:
				write_msg(Severity.Text, 'Выполнение скрипта отменено')
	gc.collect()