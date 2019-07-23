import serial
import time

#текущая дата и время
def timenow():
    return time.strftime('%Y-%m-%d|%H:%M:%S') + '|'

#функция записи лога
def add_log(log):
    a = open(fname + ' ' + timenow()[:10] + '.txt', 'a')
    a.write(timenow() + str(log))
    a.write('\n')
    a.close()
    print(timenow() + str(log))

#считывает последнюю запись в отчете
def last_log(file):
    d = open(file + ' ' + timenow()[:10] + '.txt')
    return d.readlines()[-1]
    d.close()
    
#основная функция
def send(port):
    input('Для старта работы нажмите ENTER')
    while True:
        try:
            print('Попытка подключиться к', port)
            Terminal = serial.Serial(port, 9600)
            print(port, 'подключен')
            add_log('connect')
            while True:
                data = str(Terminal.readline())[2:-5]
                if data == '1':
                    add_log('pressed')
                elif data == '0':
                    print('---')
            
        except Exception:
            print('disconnect, переподключаюсь к ', port)
            add_log('disconnect')
            time.sleep(6)
            continue

'''============================================================'''
     
fname = 'LOG'
send('com5')
