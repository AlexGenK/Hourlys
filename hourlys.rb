# гем для создания xlsx таблиц
require 'axlsx'

# --------------------------------------------------класс- электросчетчик-------------------------------------------------------
class Elmeter

	# атрибуты - наименование счетчика, предприятие владелец и почасовка	
	attr_reader :elmname
	attr_reader :hostelm
	attr_reader :hourly

	# конструктор считывает параметры счетчика из заданного файла, подменяя имя владельца
	# значением из хеша алиасов, если оно существует
	def initialize(filename, alias_name)

		@elmname=""
		@hostelm=""
		@hourly=[]

		# открываем файл в кодировке 1251
		input=File.open(filename, "r:windows-1251") 
		inputline=""

		# считываем третью строку файла где записаны номер счетчика и его владелец
		3.times {inputline=input.gets.strip}

		# определяем позицию наименования счетчика по последнему вхождению в строку подстроки с двумя цифрами и точкой после них
		# сразу после наименования счетчика следует наименование владельца
		searchpos=inputline.rindex(/\d{2}./)

		# извлекаем наименование счетчика 
		@elmname=inputline[0..searchpos+2]

		# извлекаем наименование владельца
		nam=inputline[searchpos+3..inputline.size]

		# если имя существует в хеше алиасов, заменяем имя алиасом
		if alias_name[nam]
			@hostelm=alias_name[nam]
		else
			@hostelm=nam
		end

		# пропускаем 5 строк
		6.times {inputline=input.gets.strip}

		# начинается исчисление контрольных сумм
		# по строкам
		ctrsum1=0
		# общая сумма в отдельной ячейке
		ctrsum2=0
		# по столбцам
		ctrsum3=0

		# пока не встретился конец таблицы, берем первые цифры в строке и суммируем в ctrsum1
		while inputline!="-"*238 do
			ctrsum1+=inputline[12..21].to_i
			inputline = input.gets.strip
		end

		# считываем строку с итогами таблицы. Первый элемент в ctrsum2
		a = input.gets.strip.split("!")
		a.map! {|i| i.to_i}
		ctrsum2=a[1]

		# остальные элементы - итоги по столбцам суммируем в ctrsum3
		a=a.drop(2)
		0.upto(a.size-1) {|i| ctrsum3+=a[i]}

		# высчитываем абсолютные отклонения контрольных сумм друг от друга
		otkl1=ctrsum1-ctrsum2
		otkl2=ctrsum1-ctrsum3
		otkl3=ctrsum2-ctrsum3

		input.close

		# если отклонения невелики - присваиваем нашу почасовку, иначе выводим сообщение об ошибке
		# и останавливаем скрипт
		if (otkl1.abs<10)&&(otkl2.abs<10)&&(otkl3.abs<10)
			@hourly=a
		else
			puts "Script stop! File #{filename} include incorrect data."
			exit
		end
	end

end

# --------------------------------класс - пользователь электросчетчиков-----------------------------------------------
class Eluser

	# аттрибуты - наименвоание пользователя, почасовка в абсолютных величинах, 
	# почасовка в относительных величинах, хеш со счетчиками ключ которого - наименование счетчика
	attr_reader :elusername
	attr_reader :hourlyabs
	attr_reader :hourlyperc
	attr_reader :elmeters

	# инициалируется по объекту - счетчику
	def initialize(elmeter)
		@elusername=elmeter.hostelm
		@hourlyabs=elmeter.hourly
		perc_redefine
		@elmeters={elmeter.elmname => elmeter}
	end

	# добавление счетчика, суммироание почасовок
	def add(elmeter)
		0.upto(23) {|i| @hourlyabs[i]+=elmeter.hourly[i]}
		perc_redefine
		@elmeters[elmeter.elmname] = elmeter
	end

	private

	# внутренняя процедура, которая пересчитывает почасовки в абсолютных величинах
	# в почасовки в процентах/относительных величинах
	def perc_redefine
		sum=0
		@hourlyperc=[]
		@hourlyabs.each {|i| sum+=i}
		0.upto(23) {|i| @hourlyperc[i]=100*(@hourlyabs[i].to_f/sum)}
	end

end

# ---------------------------------------------тело скрипта---------------------------------------------------------

# заполняем хэш, где хранятся алиасы для имен владельцев счетчиков
alias_name={}
input=File.open("hourlys.cfg", "r:windows-1251")
while line=input.gets
	a=line.split("|")
	alias_name[a[0]]=a[1] 
end


# получаем список текстовых файлов в текущей директории
a = Dir.entries(Dir.pwd)
a.select! {|i| File.extname(i)==".txt"}

# этот хеш хранит всех пользователй электросчетчиков. ключ - имя пользователя
h={}

# перебираем все текстовые файлы в директории. каждый файл - новый объект-счетчик
# если такой владелец счетчика уже есть в хеше - то добаляем владельцу счетчик,
# если нет, создаем нового владельца данного счетчика
a.each do |filename|
	elm=Elmeter.new(filename, alias_name)
	if h[elm.hostelm]
		h[elm.hostelm].add(elm)
	else
		h[elm.hostelm]=Eluser.new(elm)
	end
end

# формируем xlsx таблицу для записи значений
Axlsx::Package.new do |p|
  	p.workbook do |wb|  

  		# определяем стили записываемых данных
  		styles = wb.styles
  		name_style = styles.add_style :sz => 10
  		perc_style = styles.add_style :sz => 10, :num_fmt => 2

  		# добавляем строки в таблицу
	    wb.add_worksheet do |sheet|
	      	sheet.add_row
	      	h.each_value do |pred|

	      		# массив в котором формируются значения для ячеек строки
	      		a=[" "]
				a << pred.elusername.encode('utf-8')
				pred.hourlyperc.each {|hour| a << hour}
				sheet.add_row a, :style => [nil,name_style]+[perc_style]*24
			end
			sheet.column_widths 3.29, 32.71, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43, 8.43
	    end  
  	end

  	# запись в файл
  	p.serialize 'output.xlsx' 
end