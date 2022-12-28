import pandas as pd
import win32com.client as win32
import datetime as dt
# import xlwings as xw
import numpy as np

#%%Сумма текущего остатка сети // Запускаем по пн
now = dt.datetime.now()

def last_day_of_month(any_day):
    # Эта функция, которая на вход получает дату в формате (год, месяц, день),
    # а на выходе выдаёт последний день месяца, которому пренадлежит входящая дата.
    next_month = any_day.replace(day=28) + dt.timedelta(days=4)
    return next_month - dt.timedelta(days=next_month.day)



if dt.datetime.weekday(now) == 0:
    
    ALLSTOCK = pd.read_csv(r'\\Mdfsv\управление заказами\OUT_\ALLSTOCK\ALLSTOCK.csv', sep=';', encoding='1251',
                            usecols=['MS6 code', 'Art. code', 'Art. desc', 'Ext. supplier','Site', 'CurrentStock', 'StockWay', 'StockOrder', 'PeriodDays', 'DeliveryDays'])
    
    Tov_kl_od = pd.read_excel(r'\\Mdfsv\управление заказами\OUT_\ALLSTOCK\Товарный классификатор (ОД).xlsx',
                            usecols=['{Код подкатегории}', 'Характер спроса', 'Первый месяц сезонного спроса', 'Последний месяц сезонного спроса'])
    
    ALLSTOCK_IHS = pd.merge(ALLSTOCK, Tov_kl_od, left_on=['MS6 code'], right_on=['{Код подкатегории}'])
    
    ALLSTOCK_IHS.drop(labels=['MS6 code', '{Код подкатегории}'], axis=1, inplace=True)
    
    ALLSTOCK_IHS =  ALLSTOCK_IHS[ ALLSTOCK_IHS['Характер спроса'] == 'Ивентовый характер спроса']
    
    Ostatok = ALLSTOCK_IHS[['Art. code', 'Art. desc', 'CurrentStock', 'StockWay']]
    
    Ostatok = Ostatok.groupby(by='Art. code').sum()
    
    
#%%Находим INO на 101 !Здесь надо из номера месяца получить первый день ИХС и последний день ИХС
   
    INO = ALLSTOCK_IHS[['Art. code', 'Site', 'StockOrder','Ext. supplier', 'PeriodDays', 'DeliveryDays', 'Первый месяц сезонного спроса', 'Последний месяц сезонного спроса']]

    INO = ALLSTOCK_IHS[ALLSTOCK_IHS['Site'] == 101]
    
    
    
    INO['Последний месяц сезонного спроса'].fillna(0, inplace = True)
    
    INO['Последний месяц сезонного спроса'] = INO['Последний месяц сезонного спроса'].astype(int)  
    
    INO['Первый месяц сезонного спроса'].fillna(0, inplace = True)
    
    INO['Первый месяц сезонного спроса'] = INO['Первый месяц сезонного спроса'].astype(int)  
    
    conditions=[
        (INO['Последний месяц сезонного спроса'] == 0),
        (INO['Последний месяц сезонного спроса'] >= dt.datetime.now().month),
        (INO['Последний месяц сезонного спроса'] < dt.datetime.now().month)
        ]
    choices=[
        dt.date(1900, 1, 1),
        str(dt.datetime.now().year) + '-' + INO['Последний месяц сезонного спроса'].astype(str) + '-' + '1',
        str(dt.datetime.now().year + 1) + '-' + INO['Последний месяц сезонного спроса'].astype(str) + '-' + '1'
        ]
    INO['Конец СХС(ИХС)'] = np.select(conditions, choices)
    
    INO['Конец СХС(ИХС)'] = pd.to_datetime(INO['Конец СХС(ИХС)'])
    
    INO['Конец СХС(ИХС)'] = INO['Конец СХС(ИХС)'].apply(last_day_of_month)
    
    conditions=[INO['Конец СХС(ИХС)'].dt.year == 1900]
    choices=['']
    INO['Конец СХС(ИХС)'] = np.select(conditions, choices, default = INO['Конец СХС(ИХС)'].dt.date)

    INO = INO[INO['Конец СХС(ИХС)'] != '']

    INO['Конец СХС(ИХС)'] = pd.to_datetime(INO['Конец СХС(ИХС)'])             
                          
    INO['Начало СХС(ИХС)'] = np.where(
        INO['Первый месяц сезонного спроса'] <= INO['Последний месяц сезонного спроса'],
        INO['Конец СХС(ИХС)'].dt.year.astype(str) + '-' + INO['Первый месяц сезонного спроса'].astype(str) + '-' + '1',
        (INO['Конец СХС(ИХС)'].dt.year - 1).astype(str) + '-' + INO['Первый месяц сезонного спроса'].astype(str) + '-' + '1'
        )

    
    INO['Начало СХС(ИХС)'] = pd.to_datetime(INO['Начало СХС(ИХС)'])
    
    INO.drop(labels=['Первый месяц сезонного спроса', 'Последний месяц сезонного спроса'], axis=1, inplace=True)
    
 
    
                                                                     
#%%Сумма текущего остатка сети + в пути на РЦ по товарам ИХС
    Full_ostatok = pd.merge(Ostatok, INO, left_on=['Art. code'], right_on=['Art. code'], suffixes=('', '_y'))
   
    Full_ostatok['Остаток сети + в пути на РЦ'] = Full_ostatok[['CurrentStock', 'StockWay', 'StockOrder']].sum(axis=1)
    
    Full_ostatok.drop(labels=['Site', 'CurrentStock', 'StockWay', 'StockOrder'], axis=1, inplace=True)
    
#%%Находим ПП на ИХС
    Go_base = pd.read_excel(r"N:\ОУЗ\!Группа по управлению прогнозами продаж\!ПРОГНОЗЫ\ПП_ИХС\!GOLD_ORD _ПП ИХС.xlsx", "БАЗА")
    Go_base['План продаж на ИХС'] = Go_base.sum(axis=1)
    
    pp = Go_base[['Код Голд', 'План продаж на ИХС']]
    print(pp.head())

#%%Соединяем текущий остаток и пп на ихс,выбираем товары у которых остаток < пп

    Tov_ihs = pd.merge(Full_ostatok, pp, left_on=['Art. code'], right_on=['Код Голд'])
    
    Tov_ihs.drop(labels=['Код Голд'], axis=1, inplace=True)
    
    Tov_ihs["Кол-во дней продаж исх."] = Tov_ihs['Конец СХС(ИХС)'] - Tov_ihs['Начало СХС(ИХС)']

    Tov_ihs["Кол-во дней продаж исх."] = Tov_ihs["Кол-во дней продаж исх."].astype(np.int64)/86400000000000                                                                    
                                                                        

    Tov_ihs["Кол-во дней продаж ост."] = np.where(
        now > Tov_ihs['Начало СХС(ИХС)'],
        Tov_ihs['Конец СХС(ИХС)'] - now + pd.Timedelta(days = 1),
        Tov_ihs['Конец СХС(ИХС)'] - Tov_ihs['Начало СХС(ИХС)']
        )

    Tov_ihs['Кол-во дней продаж ост.'] = Tov_ihs['Кол-во дней продаж ост.'].astype(np.int64)/86400000000000

    Tov_ihs['Кол-во дней продаж ост.'] = Tov_ihs['Кол-во дней продаж ост.'].astype(int)

    Tov_ihs['План продаж на ИХС'] = (Tov_ihs['План продаж на ИХС'] / Tov_ihs['Кол-во дней продаж исх.'] * Tov_ihs['Кол-во дней продаж ост.']).astype(int)
    
    Tov_ihs.drop(labels=['Кол-во дней продаж исх.', 'Кол-во дней продаж ост.'], axis=1, inplace=True)

    Tov_ihs = Tov_ihs[(Tov_ihs['Остаток сети + в пути на РЦ'] < Tov_ihs['План продаж на ИХС'] * 0.95)]

#%%Фильтруем нужные товары, исходя из дат
    l_post = Tov_ihs['DeliveryDays'].fillna(0).reset_index(drop=True)

    t_post = Tov_ihs['PeriodDays'].fillna(0).reset_index(drop=True)

    date3 = [] 
    for i in range(Tov_ihs.shape[0]):
        date3.append(now+pd.Timedelta(days = t_post[i]) + pd.Timedelta(days = l_post[i] - 1)) 

    Tov_ihs['Дата прихода след. заказа'] = date3

    
    Tov_ihs = Tov_ihs[(Tov_ihs['Дата прихода след. заказа'] >= Tov_ihs['Начало СХС(ИХС)']-pd.Timedelta(days = 44)) & (now <= Tov_ihs['Конец СХС(ИХС)']) & (Tov_ihs['Ext. supplier'] != "-1")]

    Muz = pd.read_csv(r'\\Mdfsv\управление заказами\OUT_\SSBI\SSBI Товары с характеристиками.csv', sep=';', encoding='1251',
                            usecols=['Поставщик', 'МУЗ'])


    Muz = Muz.drop_duplicates(subset = 'Поставщик', keep="last").reset_index(drop=True)

    Tov_ihs = pd.merge(Tov_ihs, Muz, left_on=['Ext. supplier'], right_on=['Поставщик'])

    Tov_ihs.drop(labels=['PeriodDays', 'DeliveryDays', 'Дата прихода след. заказа'], axis=1, inplace=True)
    
    Tov_ihs.rename(columns = {'Art. code' : 'ТОВАР', 'Art. desc' : 'ОПИСАНИЕ', 'Ext. supplier' : 'ПОСТАВЩИК'}, inplace = True)

    Tov_ihs = Tov_ihs.reindex(columns=['ТОВАР', 'ОПИСАНИЕ', 'ПОСТАВЩИК', 'МУЗ', 'Начало СХС(ИХС)', 'Конец СХС(ИХС)', 'Остаток сети + в пути на РЦ', 'План продаж на ИХС'])

    
    Tov_ihs['Конец СХС(ИХС)'] = pd.to_datetime(Tov_ihs['Конец СХС(ИХС)']).dt.date
    
    Tov_ihs['Начало СХС(ИХС)'] = pd.to_datetime(Tov_ihs['Начало СХС(ИХС)']).dt.date
    
    
    Tov_ihs_old =  pd.read_csv(r'N:\ОУЗ\!Группа РЦ\Контроль заказа ИХС.csv', sep=';', encoding='1251')
    
    Tov_ihs_new = pd.concat([Tov_ihs_old, Tov_ihs[~Tov_ihs.ТОВАР.isin(Tov_ihs_old.ТОВАР)]])
    
    Tov_ihs_old_solve = Tov_ihs_old.loc[:,['ТОВАР','Решение']]
    
    
    
    Tov_ihs = pd.merge(Tov_ihs, Tov_ihs_old_solve, on=['ТОВАР'], how='left')
    
   
    
    Tov_ihs.to_csv(r'N:\ОУЗ\Пищик\3.Рабочая\Контроль заказа ИХС.csv', sep=';', encoding='1251', index = False)
#%%Сохраняем файл и отправляем по почте
    Tov_ihs_new.to_csv(r'N:\ОУЗ\!Группа РЦ\Контроль заказа ИХС.csv', sep=';', encoding='1251', index = False)

    outlook = win32.Dispatch('outlook.application')
    emails = ['sh_Inventory_Management_Division@boobl-goom.ru']
    for email in emails:
        mail = outlook.CreateItem(0)    
        mail.To = email    
        mail.Subject = 'Список товаров ИХС с остатком меньше ПП'    
        mail.HTMLBody = 'Коллеги, добрый день! <br/><br/>Во вложении список товаров ИХС, у которых текущий остаток меньше чем ПП на ИХС. <br/>Необходимо запланировать обеспечение остатками продаж ИХС. <br/>Результат необходимо внести в колонку "Решение" в файле "Контроль заказа ИХС": N:\ОУЗ\!Группа РЦ <br/><br/>ПОЖАЛУЙСТА, НЕ ОТВЕЧАЙТЕ. ЭТО АВТОМАТИЧЕСКАЯ РАССЫЛКА.' 
        attachment  = r"N:\ОУЗ\Со\3. Рабочая\Контроль заказа ИХС.csv"
        mail.Attachments.Add(attachment)
        mail.Send()



