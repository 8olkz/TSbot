import aiogram
import dotenv 
import telebot
import sqlite3
import re
from dotenv import load_dotenv
from aiogram import Bot,Dispatcher,executor,types
from aiogram.types import ReplyKeyboardMarkup, InlineKeyboardMarkup,InlineKeyboardButton, ContentType
from aiogram.types import ParseMode
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import StatesGroup,State
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from datetime import datetime
import os
import openpyxl
from pathlib import Path
import time
import asyncio



load_dotenv()
bot = Bot(os.getenv('Token'),parse_mode=types.ParseMode.HTML)
storage = MemoryStorage()
dp = Dispatcher(bot,storage=storage)
sqlite_connect = sqlite3.connect('ts.db')
cursos = sqlite_connect.cursor()


class starts(StatesGroup):
    namesotr = State()
    lastname = State()
    otchestwo = State()

class zp(StatesGroup):
     vibor =State()
     vibordata = State()
     vibordata2 = State()

class zp1(StatesGroup):
     vibor = State()
     viborobj=State()
     vibordata = State()
     vibordata2= State()



class reports(StatesGroup):
     name = State()
     namedata = State()
     namevaluse = State()
     sotr=State()
     yes =State()
  
class newobject(StatesGroup):
    nameobj	 = State()
    
class prosotr_smeti(StatesGroup):
     name = State()


main = ReplyKeyboardMarkup(resize_keyboard=True)
main.add('Главное меню').add('Создать отчёт').add('/cancel').add('Просмотр сметы')

adminmain=ReplyKeyboardMarkup(resize_keyboard=True)
adminmain.add('Главное меню').add('Админ-панель').add('Создать отчёт').add('/cancel').add('Просмотр сметы')

adminpanel= ReplyKeyboardMarkup(resize_keyboard=True)
adminpanel.add("Добавить объект").add("Отчёт по сотрудникам").add('Отчёт по объектам').add("Главное меню").add('/cancel')

register= ReplyKeyboardMarkup(resize_keyboard=True)
register.add("/Регистрация")


mainreport = ReplyKeyboardMarkup(resize_keyboard=True)
mainreport.add("/cancel").add("Главное меню").add("0")



global a
@dp.message_handler(commands=['main'])
async def ksksksk(message: types.Message):
     await message.answer('Вы попали в главное меню',reply_markup=main)

@dp.message_handler(state='*',commands='cancel')
async def cancel_handler(message: types.Message, state: FSMContext):
     await state.finish()
     await message.answer("Создание отчёта прервано",reply_markup=main)
           
@dp.message_handler(text = 'Просмотр сметы')
async def cmd_smeta(message: types.Message):
     await message.answer("Введите название объекта")
     cursos.execute(f"""Select name from sqlite_master where type='table'""")
     result = cursos.fetchall()
     resulttext =''
     for row in result:
          resulttext+=f"{row[0]}\n"
     await message.answer(f"<pre>{resulttext}</pre>",parse_mode=ParseMode.HTML)
     await prosotr_smeti.name.set()


@dp.message_handler(state=prosotr_smeti.name)
async def keke(message: types.Message, state: FSMContext):
     name = message.text
     cursos.execute(F"""Select id,name,value from {name}""")
     result= cursos.fetchall()
     resulttext = ''
     if result:
          for row in result:
               resulttext+=f"{row[0]}\t\t{row[1]}\t\t{row[2]}\n"
          await message.answer(f"<pre>{resulttext}</pre>",parse_mode=ParseMode.HTML,reply_markup=main)
          await state.finish()
     else:
          await message.answer("Такой таблицы нету")



@dp.message_handler(text ='Отчёт по объектам')
async def cmd_otche5t(message: types.Message):
     await zp1.vibor.set()
     await message.answer("Выберите сотрудника ")
     cursos.execute("""select * from sotr""")
     result1= cursos.fetchall()
     text1 = ''
     for row in result1:
          text1+= f"{row[0]} | {row[1]} | {row[2]} | {row[3]}\n"

     await message.answer(text1)

@dp.message_handler(state= zp1.vibor)
async def cmd_otche4t(message: types.Message, state: FSMContext):
     cursos.execute(f"select id from sotr")
     stateidsotr = message.text
     namesotr = ''
     result = cursos.fetchall()
     for row in result:
          if int(stateidsotr) > row[0]:
               message.answer("Такого сотрудника нету")
          else: 
               cursos.execute(f"select lastname,name,otchestwo from sotr where id= {stateidsotr[0]}")
               resultsotr = cursos.fetchall()
     for row in resultsotr:
          namesotr+= f"{row[0]} {row[1]} {row[2]}"
          await message.answer(f"Выбранный сотрудник {namesotr}")
          await state.update_data(sotr = namesotr)
    
     await message.answer("Выбери объект :",reply_markup = mainreport)
     cursos.execute(f"""Select name from sqlite_master where type='table'""")
     result2 = cursos.fetchall()
     resulttext1= ''
     for row in result2:
          resulttext1+=f"{row[0]}\n"
     await message.answer(resulttext1)



     await zp1.next()


@dp.message_handler(state= zp1.viborobj)
async def kek3(message: types.Message, state: FSMContext):
     await state.update_data(nameobject= message.text)
     stateidnumber = message.text
     namevibraniyobject=''
     cursos.execute(f"""Select name from sqlite_master where type='table' and name = '{stateidnumber}'""")
     idnumber = cursos.fetchall()
     for row in idnumber:
          if stateidnumber:
               await state.update_data(name = message.text)
               await message.answer(f"Выбранный объект:   <pre>{stateidnumber}</pre>",parse_mode=types.ParseMode.HTML)
               await message.answer("Введите первое число в формате ГГГГ-ММ-ДД")
               await zp1.next()

          else:
               await message.answer("Такого объекта нету")
   


@dp.message_handler(state= zp1.vibordata)
async def dmcotchet2(message: types.Message, state: FSMContext):
     await state.update_data(data1 = message.text)
     await message.answer("Дата успешно выбрана")
     await message.answer("Введите второе число в формате ГГГГ-ММ-ДД")
     await zp1.next()

@dp.message_handler(state= zp1.vibordata2)
async def dmcotchet1(message: types.Message, state: FSMContext):
     datatwo = message.text
     async with state.proxy() as date:
          await state.update_data(data2=message.text)
     cursos.execute(f"""SELECT sotr,nameobj,fielt,sum(volue),priceone from trash WHERE  date BETWEEN ('{date['data1']}') and ('{datatwo}') and nameobj='{date['name']}' group by fielt ORDER by fielt""")
     result= cursos.fetchall()
     resulttext =''
     for row in result:
          resulttext += (f"{row[0]} | {row[1]} | {row[2]} | {row[3]} | {row[4]}\n")
          money = row[3]*row[4]
     await message.answer(f"<pre>({resulttext})</pre> <pre>Сумма Р: {money}</pre>",parse_mode=ParseMode.HTML)


     await state.finish()
     await message.answer("Отчёт успешно создан",reply_markup=adminpanel)
############################################################################
@dp.message_handler(text ='Отчёт по сотрудникам')
async def cmd_otchet(message: types.Message):
     await zp.vibor.set()
     await message.answer("Выберите сотрудника ")
     cursos.execute("""select * from sotr""")
     result1= cursos.fetchall()
     text1 = ''
     for row in result1:
          text1+= f"{row[0]} | {row[1]} | {row[2]} | {row[3]}\n"

     await message.answer(text1)

@dp.message_handler(state= zp.vibor)
async def cmd_otchet(message: types.Message, state: FSMContext):
     cursos.execute(f"select id from sotr")
     stateidsotr = message.text
     namesotr = ''
     result = cursos.fetchall()
     for row in result:
          if int(stateidsotr) > row[0]:
               message.answer("Такого сотрудника нету")
          else: 
               cursos.execute(f"select lastname,name,otchestwo from sotr where id= {stateidsotr[0]}")
               resultsotr = cursos.fetchall()
     for row in resultsotr:
          namesotr+= f"{row[0]} {row[1]} {row[2]}"
          await message.answer(f"Выбранный сотрудник {namesotr}")
          await state.update_data(sotr = namesotr)
     await message.answer("Введите первое число в формате ГГГГ-MM-ДД")
     await zp.next()

@dp.message_handler(state= zp.vibordata)
async def dmcotchet(message: types.Message, state: FSMContext):
     await state.update_data(data1 = message.text)
     await message.answer("Дата успешно выбрана")
     await message.answer("Введите второе число в формате ГГГГ-ММ-ДД")
     await zp.next()

@dp.message_handler(state= zp.vibordata2)
async def dmcotchet(message: types.Message, state: FSMContext):
     await state.update_data(data2 = message.text)
     await message.answer("Дата успешно выбрана")
     data = await state.get_data()
     cursos.execute(f"SELECT sotr,nameobj,fielt,sum(volue),priceone from trash WHERE  date BETWEEN ('{data['data1']}') and ('{data['data2']}')and sotr = '{data['sotr']}' group by fielt ORDER by fielt")
     result1= cursos.fetchall()
     resulttext=''
     for row in result1:
          resulttext = f"{row[0]} | {row[1]} | {row[2]} | {row[3]} | {row[4]}"
          money = row[3]*row[4]
          await message.answer(f"<pre>{resulttext}</pre>\n <pre>Сумма Р: {money}</pre>",parse_mode=ParseMode.HTML)
          money=0
          
          
          
          
     await state.finish()
     await message.answer("Отчёт успешно создан",reply_markup=adminpanel)



#######################################################################


@dp.message_handler(text='Админ-панель')
async def cmd_start(message: types.Message):
    await message.answer(f'Вы попали в админ панель ',reply_markup=adminpanel)

#######################################################################

@dp.message_handler(commands=['start'])
async def noname(message: types.Message):
     await message.answer('Добро пожаловать', reply_markup=main) 


#######################################################################



@dp.message_handler(text='Главное меню')
async def cmd_start(message: types.Message):
    await message.answer(f'Вы попали в главное меню ',reply_markup=main)
    if message.from_user.id==int(os.getenv('ADMIN_ID')):
        await message.answer(f'Осторожно не поламай всё', reply_markup=adminmain)
#######################################################################


#######################################################################

@dp.message_handler(text='Создать отчёт')
async def create_report(message: types.Message.clean, state: FSMContext)-> None:
    await reports.name.set()
    await message.reply("Выбери объект :",reply_markup = mainreport)
    cursos.execute("""Select name from sqlite_master where type='table' and name not in('sotr','sqlite_sequence','trash')""")
    result2 = cursos.fetchall()
    resulttext1= ''
    for row in result2:
        resulttext1=f"{row[0]}\n"
        await message.answer(resulttext1)
    

@dp.message_handler(state=reports.name)
async def process_name(message: types.Message, state: FSMContext):
     name1 = message.text
     cursos.execute(f"""Select name from sqlite_master where type='table' and name = '{name1}'""")
     result = cursos.fetchall()
     if len(result) == 0:
          await message.answer("такого объекта нету")
     else:
          async with state.proxy() as date:
               await state.update_data(name = message.text)
          await message.answer(f"Выбранна таблица {name1}") #Название таблицы
          await reports.namedata.set()
          await message.answer("Выберите ячейку в таблице")
          cursos.execute(f"""Select name from {name1}""")
          result1 = cursos.fetchall()
          resulttext=''
          for row in result1:
               resulttext=f"{row[0]}\n"
               await asyncio.sleep(1)
               await message.answer(f"<pre>{resulttext}</pre>",parse_mode=types.ParseMode.HTML)
               

@dp.message_handler(state=reports.namedata)
async def process_name(message: types.Message, state: FSMContext):
     field = message.text
     async with state.proxy() as date:

          cursos.execute(f"""select name from {date['name']} where name='{field}'""")
     result= cursos.fetchall()
     if result:
          await message.answer(f"Выбранно поле <pre>{field}</pre>",parse_mode=types.ParseMode.HTML)
          await state.update_data(fielt = message.text)
          cursos.execute(f"""select priseasone from {date['name']} where name ='{field}'""")
          result1 = cursos.fetchall()
          kek=''
          for row in result1:
                    kek=row[0]
                    await state.update_data(priceone = kek)
          await message.answer("Введите объём работы")
          await reports.namevaluse.set()
     else:
          await message.answer("Такой ячейки нету")


@dp.message_handler(state=reports.namevaluse)
async def process_value(message:types.Message, state: FSMContext):
     volume = message.text
     async with state.proxy() as date:
          await state.update_data(volume=message.text)
     await message.answer(f"Выберите сотрудников: \n"
                          f"Пример: 1,2,3,4 ")
     cursos.execute(f"""Select id,lastname,name,otchestwo from sotr """)
     result = cursos.fetchall()
     resulttext=''
     for row in result:
          resulttext+= f"{row[0]} | {row[1]} | {row[2]} | {row[3]} \n"
     await message.answer(resulttext)
     await reports.sotr.set()

@dp.message_handler(state=reports.sotr)
async def process_sotr(message: types.Message , state: FSMContext):
     
     sotrid = message.text
     data1 = await state.get_data()
     cursos.execute(f"select id,name,lastname,otchestwo from sotr where id in ({sotrid})")
     result = cursos.fetchall()
     resulttext =''
     
     for row in result:
          resulttext+= f"{row[1]} | {row[2]} | {row[3]} \n"
     await message.answer("Выбранные сотрудники :")
     await message.answer(resulttext)
     async with state.proxy() as date:
          await state.update_data(idsotr = message.text)

     await message.answer("________________")
     await message.answer(f"Выбранный объект: <pre>{data1['name']}</pre>\n",parse_mode=types.ParseMode.HTML)
     await message.answer(f"Выбранное поле: <pre>{data1['fielt']}</pre>\n",parse_mode=types.ParseMode.HTML)
     await message.answer(f"Объём работы: <pre>{data1['volume']}</pre>\n",parse_mode=types.ParseMode.HTML)
     await message.answer(f"Цена за еденицу: <pre>{data1['priceone']}</pre>\n",parse_mode=types.ParseMode.HTML)
     await message.answer("Всё верно Да/Нет")
     await reports.yes.set()

     
     

@dp.message_handler(state=reports.yes)
async def process_yes(message: types.Message, state: FSMContext):
     data1= await state.get_data()  
     vibor= message.text        
     cursos.execute(f"SELECT count(id) from sotr where id in({data1['idsotr']})")
     kek =''
     kolichestwosotr = cursos.fetchall()
     for row2 in kolichestwosotr:
         kek+= f"{row2[0]}"
     if vibor =="Да":
          cursos.execute(f"""Update {data1['name']} set Value = Value - {data1['volume']} where name = '{data1['fielt']}'""")
          sqlite_connect.commit()
          cursos.execute(f"select id,lastname,name,otchestwo from sotr where id in ({data1['idsotr']})")
          result1= cursos.fetchall()
          vibraniesotrudniki=''
          for row1 in result1:
               vibraniesotrudniki=f"{row1[1]} {row1[2]} {row1[3]}"
               sqlite_connect.execute("insert into trash values (?,?,?,?,?,?,?)",
                                      (None,
                                      data1['name'],
                                      data1['fielt'],
                                      int(data1['volume'])/int(kek),
                                      data1['priceone'],
                                      vibraniesotrudniki,
                                      message.date.date()
                                      ),
                                      )
               sqlite_connect.commit()
               money = (float(data1['volume'])*float(data1['priceone']))/int(kek)
               await message.answer(f"Отчёт сотруднику {vibraniesotrudniki} успешно добавлен",reply_markup=main)
               await state.finish()
               
          await message.answer(f"Заработанно за сегодня {money}")
          

     elif vibor =="Нет":
          await message.answer("Создание отчёта прервано")
          await state.finish()





def create_sql_table_from_excel(file_path, sheet_name, table_name):
    # Подключение к базе данных
    conn = sqlite3.connect('ts.db')
    cursor = conn.cursor()

    # Чтение данных из Excel таблицы
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    
    header_row = sheet[1]
    columns = [cell.value for cell in header_row]
    column_types = ['TEXT' for _ in columns]  # Предполагается, что все столбцы имеют тип VARCHAR(255)
    
    # Создание SQL таблицы
    create_table_query = f"CREATE TABLE {table_name} ("
    for column, column_type in zip(columns, column_types):
        create_table_query += f"{column} {column_type}, "
    create_table_query = create_table_query.rstrip(', ')
    create_table_query += ")"
    
    cursor.execute(create_table_query)
    print("Выполнен executor")
    # Вставка данных в SQL таблицу
    insert_query = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join((['?']*len(columns)))})"
    for row in sheet.iter_rows(min_row=2,values_only=True):
        cursor.execute(insert_query, row)
    conn.commit()

@dp.message_handler(text = 'Добавить объект')
async def newobject1111(message: types.Message, state: FSMContext):
     await message.answer('Пришлите документ')

@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def unknown_message(message: types.Message, content_types=ContentType.DOCUMENT):
     if message.from_id==int(os.getenv('ADMIN_ID')):
          if document := message.document:
               await document.download(
          destination_file=f'{document.file_name}')
          doc_name = message.document.file_name
          text = doc_name
          result = text.split('.')[0]

          current_dir = os.path.dirname(os.path.abspath(__file__))
          main_py_path = os.path.join(current_dir, message.document.file_name)
          await message.answer(main_py_path)
          file_path=main_py_path
          table_name=result
          create_sql_table_from_excel(file_path,'Лист3',table_name)
          await message.answer("Таблица успешно добавлена")
     else:
          await message.answer("Недостаточно прав на добавлние таблицы")




if __name__=="__main__":
    executor.start_polling(dp, skip_updates=True)