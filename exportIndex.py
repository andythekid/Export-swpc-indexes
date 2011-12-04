#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Экспорт в excel индексов S-GEOMAG активности
# взятых с http://www.swpc.noaa.gov/ftpdir/indices/old_indices/

import re
import xlwt
import sys
import optparse

# Паттерн Ap индекса
aPattern = re.compile(r'^(\d{4} \d{2} \d{2})\s+(\d+)\s+\d\s\d\s\d\s\d\s\d\s\d\s\d\s\d\s+(\d+)\s+\d\s\d\s\d\s\d\s\d\s\d\s\d\s\d\s+(\d+)\s+\d\s\d\s\d\s\d\s\d\s\d\s\d\s\d')
aHeaders = ['Date', 'Middle', 'High', 'Estimated']
# Паттерн солнечных индексов
sPattern = re.compile(r'^(\d{4} \d{2} \d{2})\s+(\d+)\s+(\d+)')
sHeaders = ['Date', 'Radio flux', 'Sunspot number']

usage = "usage: %prog [options] arg1 arg2"
p = optparse.OptionParser(usage=usage)
p.add_option("-o", "--outfile", dest="outFile", action="store", type="string",
            help="set name of output file")
p.add_option("-m", "--mode",
            default="geo",
            help="analise mode: sun, geo. [default: %default]")
(options, args) = p.parse_args()

if len(args) == 0:
  p.error("Incorrect number of arguments")
if options.outFile == None:
  p.error("Output file not set. Use -o")
if options.mode == "geo":
  pat = aPattern
  head = aHeaders
elif options.mode == "sun":
  pat = sPattern
  head = sHeaders
else:
  p.error("Bad mode option. Must be 'sun' or 'geo'.")

# Создаём книгу
wb = xlwt.Workbook()
# Добавляем страницу
ws = wb.add_sheet('Data')
# Выставляем заголовок таблицы
for x in xrange(len(head)):
  ws.write(0, x, head[x])

# Номер строки в таблице
count = 1
# Для всех имен в списке аргументов коммандной строки, не попавших в опции
for inFile in args:
  # Пытаемся открыть
  try:
    tmpFile = open (inFile, 'r')
  except:
    sys.stderr.write('Bad file: %s\n ' % inFile)
    continue
  # Читаем все строки
  lst = tmpFile.readlines()
  # Для каждой прочитанной строки
  for string in lst:
    # Применяем паттерн
    rez = pat.search(string)
    # И если он отработал
    if rez != None:
      # Записываем дату
      ws.write(count, 0, rez.group(1))
      # И все остальные столбцы
      for x in xrange(1, len(head)):
        ws.write(count, x, int(rez.group(x+1)))
      count += 1

# сохраняем конечный результат
wb.save(options.outFile)

