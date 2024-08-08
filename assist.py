# Импортируем все необходимые модули
import speech_recognition as sr  # Модуль для распознавания речи
import pyttsx3  # Модуль для синтеза речи
import datetime  # Модуль для работы со временем
from fuzzywuzzy import fuzz  # Модуль для расчета схожести строк
import random  # Модуль для генерации случайных чисел
from pyowm import OWM  # Модуль для работы с API погоды
from pyowm.utils.config import get_default_config  # Модуль для настроек API погоды
from os import system  # Модуль для выполнения команд в терминале
import sys  # Модуль для работы с системными параметрами
import webbrowser  # Модуль для работы с веб-браузером
from psutil import virtual_memory as memory  # Модуль для получения информации о памяти
import psutil  # Модуль для получения информации о системе
import platform
import wmi
from PyQt5 import QtWidgets,QtCore
import threading
import interface

owm_token = 'd08d7842d66c305cd6fb9d0023424dc4'  # Ваш ключ с сайта open weather map

class Assistant(QtWidgets.QMainWindow,interface.Ui_MainWindow,threading.Thread):
    def __init__(self):
        # Вызываем конструктор родительского класса
        super().__init__()
        # Вызываем метод setupUi() для настройки пользовательского интерфейса
        self.setupUi(self)
        # Устанавливаем соединение между кнопкой pushButton и методом start_thread
        self.btnStart.clicked.connect(self.start_thread)
        # Устанавливаем соединение между кнопкой pushButton_2 и методом stop
        self.btnStop.clicked.connect(self.stop)
        # Инициализируем флаг working в значение False
        self.working = False

        # Глобальные переменные
        self.rec = sr.Recognizer()  # движок для распознавания речи
        self.engine = pyttsx3.init()  # движок для синтеза речи
        self.voices = self.engine.getProperty('voices')  # все доступные голоса
        self.asssistantVoice = 'Artemiy'  # голос ассистента

        # настройки
        self.myCity = "нижний новгород"
        self.names = ['томас', 'томэс', 'помощник', 'ассистент']
        self.ndels = ['ладно', 'не могла бы ты', 'пожалуйста']
        self.cmds = {
            ('сколько сейчас времени', 'который час', 'текущее время'): self.time,
            ('привет', 'добрый день', 'здравствуй'): self.hello,
            ('какая погода', 'погода', 'погода на улице', 'какая погода на улице'): self.weather,
            ('пока', 'вырубись'): self.quite,
            ('выключи компьютер', 'выруби компьютер'): self.shut,
            ('перезагрузи компьютер', 'запусти перезагрузку'): self.restart_pc,
            ('добавить задачу', 'добавить заметку', 'создай заметку', 'создай задачу'): self.task_planner,
            ('список задач', 'список заметок', 'задачи', 'заметки'): self.task_list,
            ('очисти список задач','удалить задачи','очистить заметки'):self.task_cleaner,
            ('включи музыку', 'вруби музон', 'вруби музыку', 'включи музон', 'врубай музыку'): self.music,
            ('место на диске', 'сколько памяти', 'сколько памяти на диске', 'сколько места'): self.disk_usage,
            ('информация о компьютере'): self.system_info,
            ('загруженость компьютера', 'загруженость системы', 'загруженость','состояние системы', 'какая загрузка системы', 'какая загрузка'): self.check_load,
        }

    def listen(self):
        text = ''
        with sr.Microphone() as source:  # подключаемся к микрофону
            print("Скажите что-нибудь...")  # запрашиваем текст
            self.rec.adjust_for_ambient_noise(source)  # Этот метод нужен для автоматического понижени уровня шума
            audio = self.rec.listen(source)  # сохраняем аудио
            try:
                text = self.rec.recognize_google(audio, language="ru-RU")  # распознаём текст
                text = text.lower()  # переводим в нижний регистр
            except sr.UnknownValueError:
                pass

            if text != '':
                # Создаем новый элемент списка
                item = QtWidgets.QListWidgetItem()
                # Устанавливаем выравнивание текста по правому краю
                item.setTextAlignment(QtCore.Qt.AlignRight)
                # Устанавливаем текст элемента списка в формате "ВЫ: <текст>"
                item.setText('ВЫ:' + '\n' + text)
                # Добавляем элемент в список
                self.console.addItem(item)
                # Прокручиваем список вниз, чтобы видеть новый элемент
                self.console.scrollToBottom()

            print(text)  # выводим распознанный текст
            return text

    def cleaner(self, text):
        cmd = ''
        for name in self.names:  # цикл по именам
            if text.startswith(name):  # Проверка начинается ли фраза с имени Ассистента
                cmd = text.replace(name, '').strip()  # записываем в команду текст без имени

        for i in self.ndels:  # цикл по ненужным словам
            cmd = cmd.replace(i, '').strip()  # удаляем слово
            cmd = cmd.replace('  ', ' ').strip()  # удаляем лишние пробелы

        return cmd  # возвращаем готовую команду

    def recognizer(self):
        text = self.listen()  # слушаем речь
        cmd = self.cleaner(text)  # очищаем речь от лишнего

        if cmd.startswith(('открой', 'запусти', 'зайди', 'зайди на')):
            self.opener(text)

        for tasks in self.cmds:  # цикл по командам
            for task in tasks:  # цикл по фразам
                if fuzz.ratio(task, cmd) >= 80:  # Проверка: если фразы похожи на 80%
                    self.cmds[tasks]()  # запускаем функцию
                    break

    def talk(self,speech):
        # устанавливаем параметры речи
        self.engine.setProperty('voice', 'ru')  # язык речи
        self.engine.setProperty('rate', 150)  # скорость речи
        self.engine.setProperty('volume', 0.7)  # громкость
        self.engine.setProperty('stress_marker', True)  # ударения

        # устанавливаем голос
        for voice in self.voices:
            if voice.name == self.asssistantVoice:
                self.engine.setProperty('voice', voice.id)

        if speech != '':
            # Создаем новый элемент списка
            item = QtWidgets.QListWidgetItem()
            # Устанавливаем выравнивание текста слева
            item.setTextAlignment(QtCore.Qt.AlignLeft)
            # Устанавливаем текст элемента списка в формате "ТОМАС: <речь>"
            item.setText('ТОМАС:' + '\n' + speech)
            # Добавляем элемент в список
            self.console.addItem(item)
            # Прокручиваем список вниз, чтобы видеть новый элемент
            self.console.scrollToBottom()

        self.engine.say(speech)  # передаём текст, который нужно сказать
        self.engine.runAndWait()  # запускаем озвучку
        print(speech)  # Вывод сказанного текста на экран

    def time(self):
        now = datetime.datetime.now()  # Сохраняем текущее время
        text = "Сейчас " + str(now.hour) + ":" + str(now.minute)  # формируем тект для озвучки
        self.talk(text)  # озвучиваем время

    def hello(self):
        text=['Привет, чем могу помочь?', 'Здраствуйте', 'Приветствую','Хеллоу'] #список приветсвий
        say = random.choice(text) #выбираем случайную фразу
        self.talk(say) #запускаем озвучку

    def weather(self):
        config_dict = get_default_config() #запрашиваем настройки
        config_dict['language'] = 'ru' #меняем язык
        owm = OWM(owm_token, config_dict)  # подключаемся к серверу
        mgr = owm.weather_manager()  # Инициализация owm.weather_manager()
        observation = mgr.weather_at_place(self.myCity)  # подключаемся кпогоде конкретного города

        weather = observation.weather  # запрашиваем погоду
        temp = weather.temperature('celsius')['temp']  # Узнаём температуру в градусах по цельсию
        feels= weather.temperature('celsius')['feels_like']  # Узнаём температуру "ощущается" в градусах по цельсию
        temp = round(temp)  # округляем температуру до целых чисел
        feels= round(feels) # округляем температуру до целых чисел
        status = weather.detailed_status  # Узнаём статус погоды в городе
        wind = weather.wind()['speed']  # Узнаем скорость ветра
        humidity = weather.humidity  # Узнаём Влажность

        text = "В городе " + self.myCity + " сейчас " + str(status) + "\nТемпература " + str(temp) + " градусов по цельсию" +\
               "\nВлажность составляет " + str(humidity) + "%" + "\nСкорость ветра " + str(wind) + " метров в секунду"
        self.talk(text)

    def opener(self, task):
        #словарь с сайтами
        links = {
            ('youtube', 'ютуб', 'ютюб'): 'https://youtube.com/',
            ('вк', 'вконтакте', 'контакт', 'vk'): 'https:vk.com/feed',
            ('браузер', 'интернет', 'browser'): 'https://google.com/',
            ('insta', 'instagram', 'инста', 'инсту'): 'https://www.instagram.com/',
            ('почта', 'почту', 'gmail', 'гмейл', 'гмеил', 'гмаил'): 'http://gmail.com/',
        }

        if 'и' in task:#убираем из запроса "и"
            task = task.replace('и', '').replace('  ', ' ')
        task=task.split() #разделяем запрос на отдельные слова

        for t in task: #цикл по запросу
            for vals in links: #цикл по словарю со ссылками
                for word in vals: #цикл по словам для запуска
                    if fuzz.ratio(word, t) > 75: #если слово похоже на запрос
                        webbrowser.open(links[vals]) #открываем ссылку в браузере
                        self.talk('Открываю ' + t) #сообщаем об открытии
                        break

    def quite(self):
        text=['Надеюсь мы скоро увидимся', 'Рада была помочь', 'Пока пока', 'Я отключаюсь']#список прощальных слов
        say= random.choice(text) #выбираем случайную фразу
        self.talk(say)#запускаем озвучку
        self.engine.stop() #выключаем движок озвучки
        sys.exit(0) #выключаем программу

    def shut(self):
        self.talk("Подтвердите действие!") #запрашиваем подтверждение
        text = self.listen() #слушаем ответ
        print(text)#выводим ответ
        #запускаем проверку: если ответ похож на подтвердить или подтверждаю
        if (fuzz.ratio(text, 'подтвердить') > 60) or (fuzz.ratio(text, "подтверждаю") > 60):
            self.talk('Действие подтверждено')
            system('shutdown /s /f /t 10') #выключаем компьютер
            self.quite() #ассистент попрощался
        elif fuzz.ratio(text, 'отмена') > 60:
            self.talk("Действие не подтверждено")
        else:
            self.talk("Действие не подтверждено")

    def restart_pc(self):
        self.talk("Подтвердите действие!")
        text = self.listen()
        print(text)
        if (fuzz.ratio(text, 'подтвердить') > 60) or (fuzz.ratio(text, "подтверждаю") > 60):
            self.talk('Действие подтверждено')
            system('shutdown /r /f /t 10 /c "Перезагрузка будет выполнена через 10 секунд"')
            self.quite()
        elif fuzz.ratio(text, 'отмена') > 60:
            self.talk("Действие не подтверждено")
        else:
            self.talk("Действие не подтверждено")

    def task_planner(self):
        self.talk("Что добавить в список задач?")  # спрашиваем
        task = self.listen()  # слушаем

        file = open('TODO_LIST.txt', 'a', encoding='utf-8')  # открываем файл в режиме дозаписи
        text = task + '\n'  # составляем текст
        file.write(text)  # записываем
        file.close()  # закрываем файл
        self.talk(f'Задача {task} добавлена в список задач!')  # говорим что задача добавлена

    # Метод для вывода списка задач из файла TODO_LIST.txt
    def task_list(self):
        try:
            # Открываем файл TODO_LIST.txt в режиме чтения с кодировкой UTF-8
            file = open('TODO_LIST.txt', 'r', encoding='utf-8')
            # Читаем содержимое файла в переменную tasks
            tasks = file.read()
            # Закрываем файл
            file.close()
            # Если список задач пуст, произносим фразу "Список задач пуст"
            if tasks == '':
                self.talk("Список задач пуст")
            else:
                # Создаем строку с заголовком "Список задач" и содержимым файла
                text = 'Список задач:\n' + tasks
                # Произносим строку
                self.talk(text)
        except:
            # Если возникла ошибка при открытии файла, произносим фразу "Вы не добавили ни одной задачи"
            self.talk('Вы не добавили ни одной задачи')

    # Метод для очистки списка задач в файле TODO_LIST.txt
    def task_cleaner(self):
        # Открываем файл TODO_LIST.txt в режиме записи с кодировкой UTF-8
        file = open('TODO_LIST.txt', 'w', encoding='utf-8')
        # Записываем пустую строку в файл, что очищает список задач
        file.write('')
        # Закрываем файл
        file.close()
        # Произносим фразу "Список задач очищен!"
        self.talk("Список задач очищен!")

    # Метод для воспроизведения музыки
    def music(self):
        # Список текстовых фраз, которые могут быть произнесены помощником
        text = ['Приятного прослушивания!', 'Наслаждайтесь!', 'Приятного прослушивания музыки']
        # Выбираем случайную фразу из списка и произносим ее
        say = random.choice(text)
        self.talk(say)

        # Список URL-адресов музыкальных композиций
        music_list = ['https://www.youtube.com/watch?v=IjwBMrxlOOA', 'https://www.youtube.com/watch?v=qj3-riPaHx8',
                      'https://www.youtube.com/watch?v=8-_4pilz70c', 'https://www.youtube.com/watch?v=ebfboqfPYGk']

        # Выбираем случайный URL-адрес из списка и открываем его в браузере
        music = random.choice(music_list)
        webbrowser.open(music)

    def check_load(self):
        #Получаем информацию о загруженности оперативной памяти
        mem = memory()
        percent = mem.percent
        #Округляем процент загрузки памяти
        percent=round(percent)
        #Формируем строку с информацией о загрузке памяти  и озвучиваем
        text = 'Компьютер загружен на '+str(percent) + ' процентов'
        self.talk(text)

    def disk_usage(self):
        total, used, free,percent = psutil.disk_usage('/') #получаем информацио о памяти
        #переводим байты в гигабайты
        total=total//(2**30)
        used=used // (2 ** 30)
        free=free // (2 ** 30)
        percent=round(percent) #округляем до целых
        #Составляем текст и озвучиваем его
        text = "Всего" + str(total) +" гигабайт, используется "+ str(used) + "гигабайт, свободно" + \
               str(free)+ 'Системный диск используется на ' +str(percent) +'процентов'
        self.talk(text)

    def system_info(self):
        text=''
        # Получаем информацию об ОС
        os_name = platform.system() + ' ' + platform.release()
        text += 'Операционная система: '+os_name+'\n'

        # Получаем информацию о процессоре
        processor_name = platform.processor()
        processor_cores = psutil.cpu_count(logical=False)
        processor_freq = psutil.cpu_freq().current

        text+='Процессор: '+processor_name +'\n'
        text+='Количество ядер: '+str(processor_cores)+'\n'
        text+='Частота процессора: '+ str(processor_freq) +'МГц \n'

        # Получаем информацию об оперативной памяти
        mem_info = psutil.virtual_memory()
        mem_total = round(mem_info.total / 2 ** 20)
        mem_used = round(mem_info.used / 2 ** 20)
        text+='Оперативная память: Всего '+ str(mem_total) + ' МБ, Используется '+ str(mem_used) +'МБ\n'

        # Получаем информацию о видеокарте
        w = wmi.WMI(namespace="root\cimv2")
        gpu_info = w.query("SELECT * FROM Win32_VideoController")[0]
        gpu_name = gpu_info.name
        text+='Видеокарта: '+gpu_name+'\n'

        # Получаем информацию о жестком диске
        hdd_info = psutil.disk_usage('/')
        hdd_free = round(hdd_info.free / 2 ** 30)
        text+='На жестком диске свободно '+str(hdd_free)+'ГБ\n'

        #озвучиваем
        self.talk(text)

    def stop(self):
        # Устанавливаем флаг working в значение False, чтобы остановить работу потока
        self.working = False
        # Вызываем метод quite() чтобы ассистент попрощался
        self.quite()

    def start_thread(self):
        try:
            # Вызываем метод hello() чтобы ассистент поздоровался
            self.hello()
            # Устанавливаем флаг working в значение True, чтобы запустить поток
            self.working = True
            # Создаем новый поток с целевой функцией main()
            self.thread = threading.Thread(target=self.main)
            # Запускаем поток
            self.thread.start()
        except Exception as e:
            # В случае возникновения ошибки выводим сообщение об ошибке и информацию об исключении
            print('Error:', e)
            print('Type:', type(e))
            print('Traceback:', sys.exc_info()[2])

    def main(self):
        while self.working:
            try:
                self.recognizer()
            except Exception as ex:
                print(ex)

if __name__ =='__main__':
   try:
        app=QtWidgets.QApplication(sys.argv)
        window=Assistant()
        window.show()
        sys.exit(app.exec_())
   except Exception as e:
       print('Error:', e)
       print('Type:', type(e))
       print('Traceback:', sys.exc_info()[2])
