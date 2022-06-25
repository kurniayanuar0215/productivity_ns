import mysql.connector
import pandas as pd
import os

# DB CONFIG
conn = mysql.connector.connect(
    host="10.47.150.144",
    user="sqaviewjbr",
    password="5qAv13wJbr@2019#",
    database="kpi"
)
# QUERY
q_ns = '''SELECT `tanggal`, `region`, `branch` NS,
SUM(`traffic_2g_total_erlang`) traffic_2g,
SUM(`traffic_3g_total_erlang`) traffic_3g,
SUM(`payload_2g_total_mbit`) payload_2g_mbit,
SUM(`payload_3g_total_mbit`) payload_3g_mbit,
SUM(`payload_4g_total_mbit`) payload_4g_mbit,
SUM(`traffic_total_erlang`) traffic_total,
SUM(`payload_total_mbit`) payload_total_mbit
FROM `productivity_daily_branch` 
WHERE tanggal BETWEEN "'''+date_1+'''" AND "'''+date_2+'''" AND region = "JABAR" '''+query_ns+'''
GROUP BY `tanggal`, `region`, `branch`;
'''

q_ns_site = '''SELECT `tanggal`, `siteid`, `branch`, `cluster`, `nsa` ns, `kabupaten`, `kecamatan`, CONCAT(`kabupaten`,"-",`kecamatan`) kabkec,
SUM(`traffic_2g_total_erlang`) traffic_2g,
SUM(`traffic_3g_total_erlang`) traffic_3g,
SUM(`payload_2g_total_mbit`) payload_2g_mbit,
SUM(`payload_3g_total_mbit`) payload_3g_mbit,
SUM(`payload_4g_total_mbit`) payload_4g_mbit,
SUM(`traffic_total_erlang`) traffic_total,
SUM(`payload_total_mbit`) payload_total_mbit
FROM kpi.`productivity_daily_siteid`
WHERE tanggal BETWEEN "'''+date_1+'''" AND "'''+date_2+'''" AND region = "JABAR" '''+query_ns+'''
GROUP BY `tanggal`, `siteid`;
'''


name_file = 'prod_ns_'+date_1+'_'+date_2+x+final_ns+'.xlsx'

prod_ns = pd.read_sql(q_ns, conn)
prod_ns_site = pd.read_sql(q_ns_site, conn)

with pd.ExcelWriter('''F:/KY/prodns/download/'''+name_file) as writer:
    prod_ns.to_excel(
        writer, index=False, sheet_name='PROD_NS'+x+final_ns)
    prod_ns_site.to_excel(
        writer, index=False, sheet_name='PROD_NS_SITE'+x+final_ns)


query.message.bot.sendDocument(query.message.chat.id, open(
    'F:/KY/prodns/download/'+name_file, 'rb'))
