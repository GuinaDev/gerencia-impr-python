import pdfkit
import win32api
import os
import time
from datetime import datetime
import cx_Oracle


# configure pdfkit to point to our installation of wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")

caminho = r"C:\pdf_arquivos"#caminho do arquivo a ser impresso
nome_arquivo_pdf = r"C:\pdf_arquivos\imprimir.pdf"# nome do arquivo a ser imprimido
lista_arquivos = os.listdir(caminho)

cod_destino = "IMP_FARM06" 
dt_prescricao = '19/09/23'
hora_prescricao = '07:14'

def imprimir():
    for arquivo in lista_arquivos:
        data_e_hora_atuais = datetime.now()
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print("Imprimindo OS para esta Oficina", data_e_hora_atuais, "oficina: " , cod_destino)
     
    
#-------------------------------------- 
# def inserir_os():
#     dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
#     conn = None
#     conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
#     cur = conn.cursor()
#     sql = f"""INSERT INTO impressao_os (cd_os,cd_oficina,dt_impressao) 
#                 VALUES ({prescricao},{cod_destino},sysdate)"""
#     cur.execute(sql)      
#     conn.commit()                  
#     print("salvo")   
#-------------------------------------- 
def num_prescricao():
    dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
    conn = None
    conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
    cur = conn.cursor()
    sql = f"""SELECT
    min(i.cd_impressao)
FROM
    impressao i
WHERE
    i.destino = 'IMP_FARM06'"""
    cur.execute(sql)
    results = cur.fetchmany(3) 
    for res in results:
        return res[0]
#-------------------------------------- 
prescricao = num_prescricao()
if prescricao:
        def teste():
            dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
            conn = None
            conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
            cur = conn.cursor()
           
            sql = f"""select 
    i.cd_impressao,
    i.cd_prioridade,
    i.nm_relatorio,
    i.destino,
    i.dt_impressao,
    i.dt_prevista_impressao,
    i.dt_solicitacao,
    i.titulo,
    i.tp_acao,
    i.solicitante
from impressao i 
where i.destino = 'IMP_FARM06'"""
            cur.execute(sql)
            results = cur.fetchall() 
            for res in results:
                return res
            
        ress = teste()
        s = f"""
        <html>
        <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Document</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
  
        </head>
        <body style="font-size: 15px">
        <div class=container>
        <div class="pdf">
            <h5 style="font-weight: bold; font-family: Arial; color: #000000;">Santa Casa de Misericordia de Juiz de Fora</h5>
             <p>{ress[0]}</p>
             </br>
             <p>{ress[1]}</p>
             </br>
             <p>{ress[2]}</p>
             </br>
             <p>{ress[3]}</p>
             </br>
             <p>{ress[4]}</p>
             </br>
             <p>{ress[5]}</p>
             </br>
             <p>{ress[6]}</p>
             </br>
             <p>{ress[7]}</p>
             </br>
             <p>{ress[8]}</p>
             </br>
             <p>{ress[9]}</p>
              <div style="margin-top: 25">
                            <div style=""> 
                            <p style="text-align: center; font-size: 11" font-family: Arial>__________________________________________________________________________________</p>
                            <p style="font-weight: bold; text-align: center; font-family: Arial; font-size: 16">Visto Solicitante</p>
                            <p style="text-align: center; font-size: 11" font-family: Arial>_________/_________/_________</p>
                            <p style="font-weight: bold; text-align: center; font-family: Arial; font-size: 16">Data</p>
                            </div>
                        </div>
        </div>
    
        </div>

        </body>
        </html>
"""
            
        pdfkit.from_string(s, output_path = r"C:\pdf_arquivos\imprimir.pdf")
        
        # imprimir()
    
        # inserir_os()   
        
else:
    time.sleep(3)
    data_e_hora_atuais = datetime.now()
    print("Nenhuma os para ser impressa", data_e_hora_atuais, "oficina: " , cod_destino)