import pyodbc
import pandas as pd



hozes = input()

hozes = hozes.split(',')

for hoze in hozes:
    
    server = 'SQL Server'
    database = 'MASHAGHEL'
    username = 'mash'
    password = '123456'
    sql = """\
    select distinct k.cod_hozeh as 'کدواحد مالياتي',k.k_parvand as 'کلاسه پرونده',k.sal as 'سال عملکرد',cast(m.Family as char(10)) as 'نام خانوادگي',cast(m.Namee as char(10))  as 'نام',cast(m.Father_name as char(10)) as 'نام پدر',t.s_des as 'شرح فعاليت',k.Modi_Seq as 'شماره مودي',m.Eco_No as 'کد اقتصادي',m.Cod_Meli as 'کد ملي',k.sahm as 'درصد سهم',k.bazargan_mashmol as 'مشمول اتاق بازرگاني' ,cast((case k.sherakat_hamsar when 0 then 'ندارد' when 1 then 'دارد' else 'نا معلوم' end) as char(15)) as 'شراکت همسر',cast((case k.end_vaziyat when 0 then 'فعال' when 1 then 'غير فعال' when 2 then 'متوفي' else 'نا معلوم' end) as char(20)) as 'آخرين وضعيت مودي',k.gr_vares_asli as 'گروه وراث اصلي',k.gr_vares_fari as 'گروه وراث فرعي',m.address as 'آدرس',m.cod_post as 'کد پستي' from  KMLINK_inf k left join modi_inf m  on k.modi_seq = m.modi_seq left join ghabln_inf g on k.k_parvand=g.k_parvand and g.sal=k.sal and g.cod_hozeh=k.cod_hozeh  left join tabl_inf t on g.Faliat_sharh=t.s_code and t.g_code=1  where 1=1 and  replace(k.cod_hozeh,' ','') like '%{0}%'  and  replace(m.Eco_No,' ','') like '%%' and  replace(m.Cod_Meli,' ','') like '%%' and  replace(m.Family,' ','') like '%%' and  replace(m.Namee,' ','') like '%%' and  replace(m.Modi_Seq,' ','') like '%%' order by k.cod_hozeh
    """.format(hoze)
    
    r=[]
    try:
            cnxn = pyodbc.connect(
                'DRIVER={SQL Server};SERVER=10.52.0.50\AHWAZ' + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
            
            tableResult = pd.read_sql(sql, cnxn) 
    
    except:
            print("error")
            
            
    tableResult.to_excel(r"D:\Software\mashReports\{0}.xlsx".format(hoze))
    print("done for {0}".format(hoze))        