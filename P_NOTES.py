import time

def make_note(fname):
    print('Запись простоев на линии "Поятос"', time.strftime('%Y-%m-%d %H:%M:%S'))
    while True:
        n = input('Новая заметка: ')
        timenow = time.strftime('%Y-%m-%d|%H:%M:%S')

        if n.replace(' ', '') != '':    #избавляемся от пустых и пробельных строк
            a = open(fname + ' ' + timenow[:10] + '.txt', 'a')
            last_note = timenow + '|' + n + '\n'
            a.write(last_note)          #заносим строку в файл
            a.close()
            print('Данные внесены, последняя запись:', last_note)
        else:
            print('Нет данных для записи')

make_note('ПОЯТОС Комментарии')
