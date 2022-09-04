from helpers import get_update_date
import time


def replace_last(phrase):
    
    
    strToReplace   = ','
    replacementStr = ''
# Search for the last occurrence of substring in string
    pos= phrase.rfind(strToReplace)
    if pos > -1:
        # Replace last occurrences of substring 'is' in string with 'XX'
        phrase = phrase[:pos] + replacementStr + phrase[pos + len(strToReplace): ] 
        
    return phrase

def table_name(report_type, year):
    
    if (report_type == 'ezhar'):
        table = '[testdb].[dbo].[tblGhateeSazi%s]' % year
    elif (report_type == 'tashkhis_sader_shode'):
        table = '[testdb].[dbo].[tblTashkhisSaderShode%s]' % year
    elif (report_type == 'tashkhis_eblagh_shode'):
        table = '[testdb].[dbo].[tblTashkhisEblaghShode%s]' % year
    elif (report_type == 'ghatee_sader_shode'):
        table = '[testdb].[dbo].[tblGhateeSaderShode%s]' % year
    elif (report_type == 'ghatee_eblagh_shode'):
        table = '[testdb].[dbo].[tblGhateeEblaghShode%s]' % year
        
    return table
    
def create_sql_table(table, columns):
    

            
    temp = ''
    
    for c in columns: 
        temp += '[%s] NVARCHAR(MAX) NULL,\n' % c     
       
        
       
    sql_query = """
    BEGIN TRANSACTION
    BEGIN TRY 

        IF Object_ID('%s') IS NULL
        
        CREATE TABLE %s
        (
         [ID] [int] IDENTITY(1,1) NOT NULL,
         PRIMARY KEY (ID),
         %s
                                  
       
         )

        COMMIT TRANSACTION;
        
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
    END CATCH
    """ % (table, table, temp)    
        
       
    # time.sleep(400)
    return sql_query
        
        
def insert_into(table, columns):
    temp = ''
    values= ''
    
    for c in columns: 
        temp += '[%s],' % c
        
        values += '?,'
        
    values = replace_last(values)
    temp = replace_last(temp)    
    # temp = temp.replace('[تاریخ بروزرسانی],', '[تاریخ بروزرسانی]', 1)
     
    sql_insert = """
    BEGIN TRANSACTION
    BEGIN TRY 

         INSERT INTO %s
           (
        %s         
            )
           
           VALUES
           (%s)

        COMMIT TRANSACTION;
        
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
    END CATCH
    """ % (table, temp, values)               
    
    return sql_insert
      
    
def insert_into_tblAnbareKoliHist(year):

    sql_query="""
    
        
    
INSERT INTO tblAnbareKoliHist%s
SELECT 
%s AS [تاریخ بروزرسانی]
,Base.*
,hesabrasi.[سال عملکرد]
,hesabrasi.[tashkhisSadere]
,hesabrasi.[eblaghTashkhis]
,hesabrasi.[ghateeSadere]
,hesabrasi.[eblaghGhatee]

 FROM 

(SELECT 
[کد اداره]
,[نام اداره]
,ISNULL([مالیات بر درآمد شرکت ها],0)AS[مالیات بر درآمد شرکت ها]
,ISNULL([مالیات بر درآمد مشاغل],0)AS[مالیات بر درآمد مشاغل]
,ISNULL([مالیات بر ارزش افزوده],0)AS[مالیات بر ارزش افزوده]

 FROM
(SELECT 
[کد اداره]
,[نام اداره]
,[منبع مالیاتی]
,count([منبع مالیاتی]) AS [tedad]
 FROM

(SELECT main.*
,[tblTashkhisSaderShode%s].[تاریخ صدور برگه تشخیص]
,[dbo].[tblTashkhisEblaghShode%s].[تاریخ ابلاغ تشخیص] as [تاریخ ابلاغ تشخیص]
,[dbo].[tblGhateeSaderShode%s].[تاریخ برگ قطعی صادر شده] as [تاریخ برگ قطعی]
,[dbo].[tblGhateeEblaghShode%s].[تاریخ ابلاغ برگ قطعی]
,1 as tedad
FROM
(SELECT * FROM [dbo].[tblGhateeSazi%s]
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی] IN(N'مالیات بر درآمد شرکت ها',N'مالیات بر درآمد مشاغل')
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
)
UNION
SELECT * FROM [dbo].[tblGhateeSazi%s]
WHERE [منبع مالیاتی]=N'مالیات بر ارزش افزوده' and [سال عملکرد]=1399
and [شناسه ملی / کد ملی (TIN)] IN
(SELECT [شناسه ملی / کد ملی (TIN)] FROM [dbo].[tblGhateeSazi%s] WHERE [سال عملکرد]=1399 and [منبع مالیاتی] IN(N'مالیات بر درآمد شرکت ها',N'مالیات بر درآمد مشاغل')and [نوع ریسک اظهارنامه] IN (N'اظهارنامه برآوردی صفر',N'انتخاب شده بدون اعمال معیار ریسک',N'رتبه ریسک بالا',N'مودیان مهم با ریسک بالا'))
) as main
LEFT JOIN
[dbo].[tblTashkhisSaderShode%s]
ON 
[dbo].[tblTashkhisSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisSaderShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblTashkhisEblaghShode%s]
ON
[dbo].[tblTashkhisEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisEblaghShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblGhateeSaderShode%s]
ON 
[dbo].[tblGhateeSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه]and [tblGhateeSaderShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblGhateeEblaghShode%s]
ON
[dbo].[tblGhateeEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه]and [tblGhateeEblaghShode%s].[سال عملکرد]=1399
)as a
group by 

[کد اداره]
,[نام اداره]
,[منبع مالیاتی]

)
 src
pivot
(
  sum(tedad)
  for [منبع مالیاتی] in ([مالیات بر درآمد شرکت ها],[مالیات بر درآمد مشاغل],[مالیات بر ارزش افزوده])
) piv) AS base

LEFT JOIN


(SELECT 
[کد اداره]
,[نام اداره]
,[سال عملکرد]
,count([تاریخ صدور برگه تشخیص]) AS  [tashkhisSadere]
,count([تاریخ ابلاغ تشخیص]) AS  [eblaghTashkhis]
,count([تاریخ برگ قطعی]) AS [ghateeSadere]
,count([تاریخ ابلاغ برگ قطعی]) AS [eblaghGhatee]
 FROM

(SELECT main.*
,[tblTashkhisSaderShode%s].[تاریخ صدور برگه تشخیص]
,[dbo].[tblTashkhisEblaghShode%s].[تاریخ ابلاغ تشخیص] as [تاریخ ابلاغ تشخیص]
,[dbo].[tblGhateeSaderShode%s].[تاریخ برگ قطعی صادر شده] as [تاریخ برگ قطعی]
,[dbo].[tblGhateeEblaghShode%s].[تاریخ ابلاغ برگ قطعی]
,1 as tedad
FROM
(SELECT * FROM [dbo].[tblGhateeSazi%s]
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی] IN(N'مالیات بر درآمد شرکت ها',N'مالیات بر درآمد مشاغل')
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
)
UNION
SELECT * FROM [dbo].[tblGhateeSazi%s]
WHERE [منبع مالیاتی]=N'مالیات بر ارزش افزوده' and [سال عملکرد]=1399
and [شناسه ملی / کد ملی (TIN)] IN
(SELECT [شناسه ملی / کد ملی (TIN)] FROM [dbo].[tblGhateeSazi%s] WHERE [سال عملکرد]=1399 and [منبع مالیاتی] IN(N'مالیات بر درآمد شرکت ها',N'مالیات بر درآمد مشاغل')and [نوع ریسک اظهارنامه] IN (N'اظهارنامه برآوردی صفر',N'انتخاب شده بدون اعمال معیار ریسک',N'رتبه ریسک بالا',N'مودیان مهم با ریسک بالا'))
) as main
LEFT JOIN
[dbo].[tblTashkhisSaderShode%s]
ON 
[dbo].[tblTashkhisSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisSaderShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblTashkhisEblaghShode%s]
ON
[dbo].[tblTashkhisEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisEblaghShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblGhateeSaderShode%s]
ON 
[dbo].[tblGhateeSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه]and [tblGhateeSaderShode%s].[سال عملکرد]=1399
LEFT JOIN
[dbo].[tblGhateeEblaghShode%s]
ON
[dbo].[tblGhateeEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه]and [tblGhateeEblaghShode%s].[سال عملکرد]=1399
)as a
group by 

[کد اداره]
,[نام اداره]
,[سال عملکرد]
) AS hesabrasi
ON base.[کد اداره]=Hesabrasi.[کد اداره]
    """ % (year, get_update_date(),year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year)
    return sql_query

def insert_into_tblAnbare99Mashaghel(year):
    sql_query = """
    INSERT INTO AnbareMashaghel%s
SELECT 
%s as [تاریخ بروزرسانی]
,ezhar.*
, [تعداد تشخیص-اظهارنامه برآوردی صفر]
, [تعداد تشخیص-انتخاب شده بدون اعمال معیار ریسک]
, [تعداد تشخیص-رتبه ریسک بالا]
, [تعداد تشخیص-مودیان مهم با ریسک بالا]
FROM
(SELECT 
[کد اداره]
,[نام اداره]
,[سال عملکرد]
,ISNULL([اظهارنامه برآوردی صفر],0) AS [اظهارنامه برآوردی صفر]
,ISNULL([انتخاب شده بدون اعمال معیار ریسک],0) AS [انتخاب شده بدون اعمال معیار ریسک]
,ISNULL([رتبه ریسک بالا],0) AS [رتبه ریسک بالا]
,ISNULL([مودیان مهم با ریسک بالا],0) AS [مودیان مهم با ریسک بالا]
 FROM
(SELECT [کد اداره]
,[نام اداره]
,[سال عملکرد]
,[نوع ریسک اظهارنامه] 
,count(*)as tedad
FROM [dbo].[tblGhateeSazi%s]
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی]= N'مالیات بر درآمد مشاغل'
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
)
GROUP BY [کد اداره]
,[نام اداره]
,[سال عملکرد]
,[نوع ریسک اظهارنامه] )
 src
pivot
(
  sum(tedad)
  for [نوع ریسک اظهارنامه] in ([اظهارنامه برآوردی صفر],[انتخاب شده بدون اعمال معیار ریسک],[رتبه ریسک بالا],[مودیان مهم با ریسک بالا])
)piv1)ezhar


LEFT JOIN



 (SELECT 
 [کد اداره]
,[نام اداره]
,ISNULL([اظهارنامه برآوردی صفر],0) AS [تعداد تشخیص-اظهارنامه برآوردی صفر]
,ISNULL([انتخاب شده بدون اعمال معیار ریسک],0) AS [تعداد تشخیص-انتخاب شده بدون اعمال معیار ریسک]
,ISNULL([رتبه ریسک بالا],0) AS [تعداد تشخیص-رتبه ریسک بالا]
,ISNULL([مودیان مهم با ریسک بالا],0) AS [تعداد تشخیص-مودیان مهم با ریسک بالا]
 FROM
 (SELECT 
ezhar.[کد اداره]
,ezhar.[نام اداره]
,ezhar.[نوع ریسک اظهارنامه]
,count([تاریخ صدور برگه تشخیص]) as [tedad]
FROM
(SELECT [کد اداره]
,[سال عملکرد]
,[نام اداره]
,[نوع ریسک اظهارنامه] 
,[شناسه اظهارنامه]
FROM [dbo].[tblGhateeSazi%s] 
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی]= N'مالیات بر درآمد شرکت ها'
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
))as ezhar LEFT JOIN
[dbo].[tblTashkhisSaderShode%s] ON ezhar.[شناسه اظهارنامه]=[tblTashkhisSaderShode%s].[شناسه اظهارنامه] and [tblTashkhisSaderShode%s].[سال عملکرد]=1399
group by ezhar.[کد اداره]
,ezhar.[نام اداره]
,ezhar.[نوع ریسک اظهارنامه]
)
 src
pivot
(
  sum(tedad)
  for [نوع ریسک اظهارنامه] in ([اظهارنامه برآوردی صفر],[انتخاب شده بدون اعمال معیار ریسک],[رتبه ریسک بالا],[مودیان مهم با ریسک بالا])
) piv2)tashkhis
ON ezhar.[کد اداره]=tashkhis.[کد اداره]
    """ % (year, get_update_date(), year, year, year, year, year)
    
    return sql_query


def insert_into_tblAnbare99Sherkatha(year):
    sql_query = """
    
INSERT INTO AnbareSherkatha%s
SELECT 
%s as [تاریخ بروزرسانی]
,ezhar.*
, [تعداد تشخیص-اظهارنامه برآوردی صفر]
, [تعداد تشخیص-انتخاب شده بدون اعمال معیار ریسک]
, [تعداد تشخیص-رتبه ریسک بالا]
, [تعداد تشخیص-مودیان مهم با ریسک بالا]
FROM
(SELECT 
[کد اداره]
,[نام اداره]
,[سال عملکرد]
,ISNULL([اظهارنامه برآوردی صفر],0) AS [اظهارنامه برآوردی صفر]
,ISNULL([انتخاب شده بدون اعمال معیار ریسک],0) AS [انتخاب شده بدون اعمال معیار ریسک]
,ISNULL([رتبه ریسک بالا],0) AS [رتبه ریسک بالا]
,ISNULL([مودیان مهم با ریسک بالا],0) AS [مودیان مهم با ریسک بالا]
 FROM
(SELECT [کد اداره]
,[نام اداره]
,[سال عملکرد]
,[نوع ریسک اظهارنامه] 
,count(*)as tedad
FROM [dbo].[tblGhateeSazi%s]
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی]= N'مالیات بر درآمد شرکت ها'
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
)
GROUP BY [کد اداره]
,[نام اداره]
,[سال عملکرد]
,[نوع ریسک اظهارنامه] )
 src
pivot
(
  sum(tedad)
  for [نوع ریسک اظهارنامه] in ([اظهارنامه برآوردی صفر],[انتخاب شده بدون اعمال معیار ریسک],[رتبه ریسک بالا],[مودیان مهم با ریسک بالا])
)piv1)ezhar


LEFT JOIN



 (SELECT 
 [کد اداره]
,[نام اداره]
,ISNULL([اظهارنامه برآوردی صفر],0) AS [تعداد تشخیص-اظهارنامه برآوردی صفر]
,ISNULL([انتخاب شده بدون اعمال معیار ریسک],0) AS [تعداد تشخیص-انتخاب شده بدون اعمال معیار ریسک]
,ISNULL([رتبه ریسک بالا],0) AS [تعداد تشخیص-رتبه ریسک بالا]
,ISNULL([مودیان مهم با ریسک بالا],0) AS [تعداد تشخیص-مودیان مهم با ریسک بالا]
 FROM
 (SELECT 
ezhar.[کد اداره]
,ezhar.[نام اداره]
,ezhar.[نوع ریسک اظهارنامه]
,count([تاریخ صدور برگه تشخیص]) as [tedad]
FROM
(SELECT [کد اداره]
,[نام اداره]
,[نوع ریسک اظهارنامه] 
,[شناسه اظهارنامه]
FROM [dbo].[tblGhateeSazi%s] 
WHERE 
[سال عملکرد]=1399
and [منبع مالیاتی]= N'مالیات بر درآمد شرکت ها'
and [نوع ریسک اظهارنامه] IN (
N'اظهارنامه برآوردی صفر'
,
N'انتخاب شده بدون اعمال معیار ریسک'
,
N'رتبه ریسک بالا'
,
N'مودیان مهم با ریسک بالا'
))as ezhar LEFT JOIN
[dbo].[tblTashkhisSaderShode%s] ON ezhar.[شناسه اظهارنامه]=[tblTashkhisSaderShode%s].[شناسه اظهارنامه] and [tblTashkhisSaderShode%s].[سال عملکرد]=1399
group by ezhar.[کد اداره]
,ezhar.[نام اداره]
,ezhar.[نوع ریسک اظهارنامه]
)
 src
pivot
(
  sum(tedad)
  for [نوع ریسک اظهارنامه] in ([اظهارنامه برآوردی صفر],[انتخاب شده بدون اعمال معیار ریسک],[رتبه ریسک بالا],[مودیان مهم با ریسک بالا])
) piv2)tashkhis
ON ezhar.[کد اداره]=tashkhis.[کد اداره]
    """ % (year, get_update_date(), year, year, year, year, year)
    
    return sql_query


def insert_into_tblHesabrasiArzeshAfzoode(year):
    sql_query = """
    
INSERT INTO tblHesabrasiArzeshAfzode%s
SELECT
      %s AS  [تاریخ بروزرسانی]
      ,[کد اداره]
      ,[نام اداره]
      ,[سال عملکرد]
	  ,COUNT([شناسه اظهارنامه]) AS [تعداد اظهارنامه]
	  ,count([تاریخ صدور برگه تشخیص]) AS  [tashkhisSadere]
      ,count([تاریخ ابلاغ تشخیص]) AS  [eblaghTashkhis]
      ,count([تاریخ برگ قطعی]) AS [ghateeSadere]
      ,count([تاریخ ابلاغ برگ قطعی]) AS [eblaghGhatee]

 FROM

(SELECT main.*
,[tblTashkhisSaderShode%s].[تاریخ صدور برگه تشخیص]
,[dbo].[tblTashkhisEblaghShode%s].[تاریخ ابلاغ تشخیص] as [تاریخ ابلاغ تشخیص]
,[dbo].[tblGhateeSaderShode%s].[تاریخ برگ قطعی صادر شده] as [تاریخ برگ قطعی]
,[dbo].[tblGhateeEblaghShode%s].[تاریخ ابلاغ برگ قطعی]
 FROM
(SELECT [کد اداره]
      ,[نام اداره]
      ,[سال عملکرد]
      ,[شناسه اظهارنامه]

  FROM [TestDb].[dbo].[tblGhateeSazi%s]
  where  [منبع مالیاتی]=N'مالیات بر ارزش افزوده'
  ) AS main
  LEFT JOIN
[dbo].[tblTashkhisSaderShode%s]
ON 
[dbo].[tblTashkhisSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisSaderShode%s].[سال عملکرد]=main.[سال عملکرد]
LEFT JOIN
[dbo].[tblTashkhisEblaghShode%s]
ON
[dbo].[tblTashkhisEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblTashkhisEblaghShode%s].[سال عملکرد]=main.[سال عملکرد]
LEFT JOIN
[dbo].[tblGhateeSaderShode%s]
ON 
[dbo].[tblGhateeSaderShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblGhateeSaderShode%s].[سال عملکرد]=main.[سال عملکرد]
LEFT JOIN
[dbo].[tblGhateeEblaghShode%s]
ON
[dbo].[tblGhateeEblaghShode%s].[شناسه اظهارنامه]=main.[شناسه اظهارنامه] and [tblGhateeEblaghShode%s].[سال عملکرد]=main.[سال عملکرد]
)as a
 GROUP BY 
[کد اداره]
      ,[نام اداره]
      ,[سال عملکرد]
    """ % (year, get_update_date(),year, year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year ,year)
    
    return sql_query
