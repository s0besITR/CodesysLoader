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
	Обновляет активный проект. Загружает Application, обновляет Slave104Drivers
	"""
	#Списки каналов
	global start_time
	d_keys = ['data', 'cmd']
	slave_ST = dict.fromkeys(d_keys)
	slave_EDC = dict.fromkeys(d_keys)
	
	if start_time == 0:
		start_time = datetime.now()
	
	if files != None:
		for file in files:
			if file.endswith('Application.xml'):
				import_application(file)
			elif file.endswith('EDC_CMD.iec104cmd.xml'):
				slave_EDC['cmd'] = load_channels(file)
			elif file.endswith('EDC_DATA.iec104data.xml'):
				slave_EDC['data'] = load_channels(file)
			elif file.endswith('ST_CMD.iec104cmd.xml'):
				slave_ST['cmd'] = load_channels(file)
			elif file.endswith('ST_DATA.iec104data.xml'):
				slave_ST['data'] = load_channels(file) 
			else:
				write_msg(Severity.Error, 'Неизвестный формат файла:   ' + file)
	else:
		write_msg(Severity.Text, 'Выполнение скрипта отменено')
		return False
	
	if slave_ST.get('data')!= None or slave_ST.get('cmd')!= None:
		iec104slave_mod('Slave_104_Driver', slave_ST)
		iec104_GVL('I104_GVL_TM', slave_ST)
		write_msg(Severity.Text, '----------Импорт каналов ТМ закончен----------')
	if slave_EDC.get('data')!= None or slave_EDC.get('cmd')!= None:
		iec104slave_mod('Slave_104_Driver_EDC', slave_EDC)
		iec104_GVL('I104_GVL_KK', slave_EDC)
		write_msg(Severity.Text, '----------Импорт каналов EDC закончен----------')
	dir = files[0].split('\\')
	dir.pop(-1)
	replaceModbusDevices('\\'.join(dir))
	# очистка памяти
	del slave_ST
	del slave_EDC
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
				if file.endswith('.xml'):
					xml_list.append(dir + '\\' + file)
		if len(xml_list):
			if(set_primary_prj(uso_folder)):			
				modifyActiveProject(xml_list)			
				err_count += compileActivePrj()
				write_msg(Severity.Information, 'Работа с текущим устройством завершена')			
	return err_count
	
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
	
		'bo_tb_fb' : [3, 5, 7, 31, 32, 33],
		'ep_td_fb' : [38],
		'it_tb_fb' : [15, 37],
		'me_tf_fb' : [9, 11, 13, 21, 34, 35, 36],
		'sp_tb_fb' : [1, 30]
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
	Удаляем предыдущий GVL, если был. Добавляем новый. Наполняем переменными из каналов
	"""
	# Удаляем старое	
	clearOld_GVL(gvl_name)	
	
	# Добавляем пустой новый
	app = getActiveApplication()
	if not app is None :
		new_gvl_obj = app.create_gvl(gvl_name)
		new_gvl_obj.textual_declaration.remove(0, 0, new_gvl_obj.textual_declaration.length)
		new_gvl_obj.textual_declaration.insert(0, 0, getGVLdata(channels))
		write_msg(Severity.Text, 'Добавлено:   {gvl}'.format(gvl = gvl_name))
	else:
		write_msg(Severity.Error, 'Application не найдено')
		return	
	
def clearOld_GVL(gvl_name):
	"""
	Удаляем GVL, которые назывались по старому, если они еще есть в проекте
	"""
	gvl_list = {
		'I104_GVL_TM' : 'I104_GVL_1',
		'I104_GVL_KK' : 'I104_GVL_2'
	}
	obj = getActiveDevice().find(gvl_list.get(gvl_name), True)
	if len(obj) > 0 :
		obj[0].remove()
		write_msg(Severity.Text, 'Удалено:   {gvl}'.format(gvl = gvl_list.get(gvl_name)))
	obj = getActiveDevice().find(gvl_name, True)
	if len(obj) > 0 :
		obj[0].remove()
		write_msg(Severity.Text, 'Удалено:   {gvl}'.format(gvl = gvl_name))	

def getGVLdata(channels):
	"""
	Из каналов получаем внутренний текст GVL
	"""		
	gvl_text = "VAR_GLOBAL" + '\n' + '\t//Updated: ' + datetime.now().strftime('%d.%m.%Y %H:%M:%S') + '\n'
	if(channels['data']):
		for ch in channels['data']:		
			gvl_text += '\t// ' + ch.get('Descr')[1:] + '\n\t' + ch.get('MapVarName') + ': IEC_LIB.' + ch.get('lib_type') + ';\n'
	if(channels['cmd']):	
		for ch in channels['cmd']:
			gvl_text += '\t// ' + ch.get('Descr')[1:] + '\n\t' + ch.get('MapVarName') + ': IEC_LIB.' + ch.get('lib_type') + ';\n'
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
		return True
	else:
		write_msg(Severity.Text, 'Этап замены Modbus модулей пропущен, не найдена папка modules')
		return False

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
			folder = system.ui.browse_directory_dialog("Выберите папку modules", path = main_dir)
			if folder != None:
				if replaceModbusDevices(folder.replace('\\modules', '')):
					write_msg(Severity.Information, 'Завершено!')
					write_msg(Severity.Text, 'Время выполнения: {}'.format(datetime.now() - start_time))
					system.ui.info("Выполнено успешно!")
			else:
				write_msg(Severity.Text, 'Выполнение скрипта отменено')
	gc.collect()