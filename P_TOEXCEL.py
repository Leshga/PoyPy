import xlsxwriter
import time

#Начальные данные
date = '2019-07-18'         #дата измерений
start_time = '08:00:00'     #время начала смены
end_time = '17:00:00'       #время окончания смены
min_step = 5                #шаг времени в минутах
norm_press = 1000            #норматив формовок в смену
work_duration = 8           #длительность смены в часах
state_pressed = 'pressed'
state_connect = 'connect'
state_disconnect = 'disconnect'
palets_total = 0
export_name = 'ГРАФИК прессования Поятос'
log_file = 'log ' + date + '.txt'   #имя файла с логами, дата подставляется автоматически
norm_per_step = round(norm_press*min_step/(work_duration*60), 1)


#конвертирование времени в список
def timestr_to_list(ts):
    ti = [int(i) for i in ts.split(':')]
    return ti

#конвертирование времени в строку
def timelist_to_str(tt):
    ti = [str(i) for i in tt]
    for j in range(len(ti)):
        if len(ti[j]) == 1:
            ti[j] = '0' + ti[j]
    ts = ':'.join(ti)
    return ts

#конвертирует строку времени в количество секунд
def timestr_to_sec(ts):
    ti = [int(i) for i in ts.split(':')]
    return ti[2] + ti[1]*60 + ti[0]*3600

#конвертирует секунды в строку времени
def timesec_to_str(tsec):
    ss = tsec%60
    mm = (tsec//60)%60
    hh = (tsec//3600)%60
    t = [str(i) for i in [hh, mm, ss]]
    for j in range(len(t)):
        if len(t[j]) == 1:
            t[j] = '0' + t[j]
    tstr = ':'.join(t)
    return tstr   

#добавление интервала времени для графика
def extra_time(tline, val, count=1):
    hh, mm, ss = timestr_to_list(tline)
    mmplus = (mm + val)*count%60
    hhplus = (mm + val)*count//60 + hh
    return timelist_to_str([hhplus, mmplus, ss])

#вытаскивает данные логов и возвращает вложенный список в формате [время, состояние]
def get_logs(file):
    #считываем файл
    try:
        with open(file, 'r') as n:
            llist = n.readlines()
    
        #пересобираем список и вырезаем лишнее
        for i in range(len(llist)):
            llist[i] = llist[i][llist[i].find('|')+1:]
            llist[i] = llist[i].replace('\n','').split('|')
            if len(llist[i][0]) == 8 and (llist[i][1] == state_pressed or state_connect or state_disconnect): #проверяем данные на корректность
                pass
            else:
                print('Ошибка лог-файла "' + file + '", проверьте корректность логов')
                print('Файл должен иметь идентичный формат строк')
                input('Нажмите ENTER для продолжения...')

        #фильтруем коннекты, оставляем последний лог из подряд идущих "connect"
        i = 0
        while i < len(llist)-1:
            if llist[i][1] == state_connect and llist[i+1][1] == state_connect:
                llist.remove(llist[i])
            else:
                i += 1

        #добавляем числовое значение состояния 0, 1 для подсчета поддонов
        for v in range(len(llist)):
            if llist[v][1] == state_pressed:
                llist[v].insert(1, 1)
            else:
                llist[v].insert(1, 0)

        #заменяем строковые состояния на "connect", "disconnect"
        llist[0][2] = state_connect
        llist[-1][2] = state_disconnect
        con_stack = True
        indexes = []
        for g in range(1, len(llist)-1):
            if llist[g][2] == state_connect:
                if con_stack == True:
                    indexes.append(g-1)
                else:
                    con_stack = True
            elif llist[g][2] == state_disconnect:
                con_stack = False
            else:
                llist[g][2] = state_connect

        for h in indexes:
            llist[h][2] = state_disconnect
        if len(indexes) > 0:
            print('При записи показаний были случаи некорректного завершения работы логгера.')
            print('Статистика может быть искажена. (Обрезан диапазон измерений)')
            input('Нажмите ENTER для продолжения...')
            
        print('Список сформирован')
        return llist
        
    except FileNotFoundError:
        print('Указанный лог-файл отсутствует. Введите корректные данные')
        input('Нажмите ENTER для продолжения...')

#возвращает 1 или None в зависимости от состояния лога - отображает на графике диапазоны измерений, локальная функция для assemble_palets
def log_state_to_bin(state):
    if state == state_connect:
        return None
    elif state == state_disconnect:
        return 1

#подсчет поддонов, пересборка логов по таймингу заданного диапазона, окончательная подготовка к экспорту
def assemble_palets():
    result_list = []
    tm_cnt = 0              #счетчик основного цикла
    log_cnt = 0             #счетчик объектов списка лога
    palets_local = 0        #счетчик поддонов в отрезке
    global palets_total     #счетчик поддонов за все время
    log_lines = get_logs(log_file)  #список логов
    time1 = start_time              #время предыдущего интервала в цикле

    
        
    #этот цикл определяет время начала и окончания расчета произвдительности
    while timestr_to_sec(time1) <= timestr_to_sec(end_time) and log_cnt < len(log_lines):
        time1 = extra_time(start_time, min_step, tm_cnt+0)   #переопределяем time1
        time2 = extra_time(start_time, min_step, tm_cnt+1) #время следующего интервала в цикле
        log_time = log_lines[log_cnt][0]    #время события лога
        log_press = log_lines[log_cnt][1]   #для подсчета поддонов, все логи прессования имеют значение - 1, остальные - 0
        log_state = log_lines[log_cnt][2]   #описание события лога (состояние)
        
        #если лог опережает временной отрезок, то записываем результат и сдвигаем время
        if timestr_to_sec(log_time) >= timestr_to_sec(time2):
            result_list.append([time2, palets_local, log_state_to_bin(log_state)])
            palets_local = 0
            tm_cnt += 1
            
        #если лог находится до измеряемого диапазона, то переходим к следующему
        elif timestr_to_sec(log_time) < timestr_to_sec(time1):
            log_cnt += 1

        #лог попал во временной отрезок, проверяем ответ, складываем поддоны, если состояние соответствует 'state_pressed'
        else:
            log_cnt += 1             #переходим к следующему логу
            palets_local += log_press    #складываем кол-во поддонов в интервале
            palets_total += log_press    #складываем все поддоны

    return result_list

#настройка экспорта и экспорт в эксель
def export_to_excel():
    export_file = export_name + ' ' + date + '.xlsx'
    ref_list = assemble_palets()
    nor = round(1000*5/(8*60), 1)

    for i in range(len(ref_list)):
        if type(ref_list[i][2]) == int:
            ref_list[i][2] = ref_list[i][2]*nor

    work_time = 0
    laze_time = 0
    for k in range(len(ref_list)):
        if ref_list[k][2] == 1 and ref_list[k][2] > 1:
            work_time += round(min_step/60, 1)
        else:
            laze_time += round(min_step/60, 1)
    

    cheet_values_a = '=Данные!$A$1:$A$' + str(len(ref_list))
    cheet_values_b = '=Данные!$B$1:$B$' + str(len(ref_list))
    cheet_values_c = '=Данные!$C$1:$C$' + str(len(ref_list))
    col_a = [ref_list[i][0] for i in range(len(ref_list))]
    col_b = [ref_list[i][1] for i in range(len(ref_list))]
    col_c = [ref_list[i][2] for i in range(len(ref_list))]
    
    
    
    while True:
        try:    
            #настройки листа
            workbook = xlsxwriter.Workbook(export_file)
            worksheet1 = workbook.add_worksheet('График')       #Добавление листа
            worksheet2 = workbook.add_worksheet('Данные')       #Добавление листа
            worksheet1.set_landscape()                          #Альбомная ориентация
            worksheet1.set_paper(9)                             #А4
            worksheet1.center_horizontally()                    #Центровка по горизонтали
            worksheet1.center_vertically()                      #Центровка по вертикали
            worksheet1.set_margins(0.05, 0.05, 0.05, 0.05)      #поля листа для распечатанной страницы
            worksheet1.set_header ('График')                    #заголовок напечатанной страницы и параметры
            worksheet1.hide_gridlines(1)                        #не печатать сетку
            worksheet1.print_area('A1:S35')                     #Cells A1 to H20.
            worksheet1.set_print_scale (150)                    #масштабный коэффициент для распечатанной страницы
##            worksheet.write(row, 1, '=SUM(B1:B4)')

            merge_format = workbook.add_format({'align': 'center',
                                                'bold': True,
                                                'border': 1,
                                                'size': 16,
                                                'align': 'center',
                                                'valign': 'vcenter',})
            worksheet1.merge_range('B3:R3', 'График производительности линии "Поятос" по интервалам времени', merge_format)




            worksheet1.write('B6', 'Норма формовок в смену')
            worksheet1.write('F6', norm_press)
            worksheet1.write('B7', 'Фактическое число формовок')
            worksheet1.write('F7', palets_total)
            worksheet1.write('B8', 'Норма формовок за интервал')
            worksheet1.write('F8', '=F6*5/(8*60)')
            worksheet1.write('B9', 'Расч. кол. формовок за интервал')
            worksheet1.write('F9', '=F7*5/(8*60)')
            worksheet1.write('B10', 'Нормативная длительность форм., с')
            worksheet1.write('F10', '=8*3600/F6')
            worksheet1.write('B11', 'Расчетная длительность форм., с')
            worksheet1.write('F11', '=8*3600/F7')


            worksheet1.write('M6', 'Время в работе (прессование)')
            worksheet1.write('Q6', work_time)
            worksheet1.write('M7', 'Время без формовок (простой)')
            worksheet1.write('Q7', laze_time)
            worksheet1.write('M8', 'Неучтенное время')
            worksheet1.write('Q8', '=Q9-Q7-Q6')
            worksheet1.write('M9', 'Длительность смены')
            worksheet1.write('Q9', work_duration)
            worksheet1.write('M10', 'Производит. без уч. простоев')
            worksheet1.write('Q9', '')

            worksheet1.write('M13', 'Разраб.:  А.П. Бежанов')
            worksheet1.write('M14', 'Проверил:')




                
            worksheet2.write_column('A1', col_a)
            worksheet2.write_column('B1', col_b)
            worksheet2.write_column('C1', col_c)



            norm_line = workbook.add_chart({'type': 'area'})
            norm_line.add_series({
                'name':       'диапазон \nбез измерений',    #Название окна графика
                'categories': cheet_values_a,   #значения оси Х
                'values':     cheet_values_c,   #значения оси У
                'line': {
                    'color': 'red',
                    'width': 2,
                },
                'trendline': {
                    'name':       'Норматив',    #Название окна графика
                    'type': 'linear',
                    'display_equation': True
                }
            })


                        
            graph_line = workbook.add_chart({'type': 'line'}) #тип заполненная линия
            graph_line.add_series({
                'name':       'Количество\n формовок',    #Название окна графика
                'categories': cheet_values_a,   #значения оси Х
                'values':     cheet_values_b,   #значения оси У
                'line': {
                    'display_equation': True
                }
            })



            
            graph_line.set_style(33)
            graph_line.combine(norm_line)
            graph_line.set_size({'width': 1200, 'height': 400})#размер графика
            graph_line.set_chartarea({'border': {'none': True}})#отключение границы окна графика
##            graph_line.set_legend({'none': True}) #отключает легенду
            worksheet1.insert_chart('A15', graph_line, {'x_offset': 5, 'y_offset': 5})#позиция и смещение окна графика
            workbook.close()
            print(export_file, 'успешно сгенерирован.')
            break
                        
        except PermissionError:
            #если не закрыт файл, то просим закрыть
            err_msg = 'Не могу сохранить данные, закройте файл: "' + export_file + '" и нажмите ENTER'
            input(err_msg)


    
'''==================================================='''


export_to_excel()
##print('Формовок за интервал по нормативу: ', norm_per_step)
##print('Всего формовок', palets_total)
