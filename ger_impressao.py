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

cod_oficina = 29

def imprimir():
    for arquivo in lista_arquivos:
        data_e_hora_atuais = datetime.now()
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
        print("Imprimindo OS para esta Oficina", data_e_hora_atuais, "oficina: " , cod_oficina)
     
    
#-------------------------------------- 
def inserir_os():
    dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
    conn = None
    conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
    cur = conn.cursor()
    sql = f"""INSERT INTO impressao_os (cd_os,cd_oficina,dt_impressao) 
                VALUES ({os},{cod_oficina},sysdate)"""
    cur.execute(sql)      
    conn.commit()                  
    print("salvo")   
#-------------------------------------- 
def num_os():
    dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
    conn = None
    conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
    cur = conn.cursor()
    sql = f"""SELECT 
                min(so.cd_os)
                FROM
                    dbamv.solicitacao_os so,
                    dbamv.itsolicitacao_os iso,
                    cn_sla.impressao_os io
                WHERE
                    so.cd_oficina = {cod_oficina}
                AND so.tp_situacao IN (
                    'A',
                    'S'
                    )
            and so.cd_os=iso.cd_os(+)
            and iso.cd_os is null
            and so.cd_os=io.cd_os(+)
            and io.cd_os is null
            and so.dt_pedido >= to_date('05/01/2023', 'DD-MM-YYYY')
            and so.cd_tipo_os!=36"""
    cur.execute(sql)
    results = cur.fetchmany(3) 
    for res in results:
        return res[0]
#-------------------------------------- 
while True:    
    os = num_os()
    if os is not None:
        def teste():
            dsnStr = cx_Oracle.makedsn("172.18.2.193", "1521", "prd2")
            conn = None
            conn = cx_Oracle.connect(user="cn_sla", password="%sla*", dsn=dsnStr)
            cur = conn.cursor()
           
            sql = f"""SELECT 
                    nvl( to_char(so.dt_prev_exec, 'dd/mm/yy hh24:mi'),to_char(so.dt_pedido, 'dd/mm/yy hh24:mi')) AS dt_pedido,
                    nvl((select u.nm_usuario from DBASGU.usuarios u where u.cd_usuario=so.nm_solicitante ),so.nm_solicitante) as nm_solicitante,
                    to_char(so.dt_execucao, 'dd/mm/yy hh24:mi') AS dt_execucao,
                    so.DS_OBSERVACAO as DS_OBSERVACAO,
                    SO.DS_SERVICO,
                    so.cd_os,
                    so.DS_RAMAL,
                    so.cd_tipo_os||' - '||ts.ds_tipo_os as ds_tipo_os,
                    s.cd_setor||' - '||s.nm_setor as nm_setor,
                    l.cd_localidade||' - '||l.ds_localidade as ds_localidade,
                    b.ds_plaqueta||' - '||b.ds_bem as ds_plaqueta,
                    case  when so.tp_situacao='C' then 'Concluído' 
                        when so.tp_situacao='N' then 'Não atendido'
                        when so.tp_situacao='E' then 'Concerto externo' 
                        when so.tp_situacao='M' then 'Aguardando Material' else 'Aberto' end as tp_situacao,
                    so.ds_ramal,
                    ( case  when avs.tp_status='A' then 'Aprovado' 
                        when avs.tp_status='R' then 'Reprovado' else 'Não avaliado' end ||' - '||
                    case  when avs.tp_avaliacao='E' then 'Excelente' 
                        when avs.tp_avaliacao='B' then 'Bom'
                        when avs.tp_avaliacao='R' then 'Razoável' 
                        when avs.tp_avaliacao='U' then 'Ruim' else 'Não avaliado' end )as avaliacao,
                        avs.ds_observacao as obsAvaliacao,
                    (select u.nm_usuario from dbasgu.usuarios u where u.cd_usuario=avs.cd_usuario) as avaliador,
                    nvl((select u.nm_usuario from dbasgu.usuarios u where u.cd_usuario=osec.cd_usuario),' -- ') as destimunha,
                    to_char(osec.dt_confirmacao, 'dd/mm/yy hh24:mi') as dt_conf_comp,
                    so.cd_oficina||' - '||o.ds_oficina as ds_oficina,
                    me.ds_espec,
                    to_char(avs.dt_ultima_atualizacao,'dd/mm/yy hh24:mi'),
                    so.ds_servico_geral,
                    s.cd_setor,
                    l.cd_localidade
                    
                FROM
                    dbamv.solicitacao_os so,
                    DBAMV.TIPO_OS TS,
                    dbamv.setor s,
                    dbamv.localidade l,
                    dbamv.avaliacao_servico   avs,
                    dbamv.bens b,
                    cn_sla.ordens_serv_entr_conf osec,
                    dbamv.oficina o,
                    DBAMV.manu_espec me
                WHERE
                    so.cd_tipo_os=TS.CD_TIPO_OS
                    and so.cd_setor=s.cd_setor
                    and so.cd_localidade=l.cd_localidade
                    and so.cd_bem=b.cd_bem(+)
                    and so.cd_os=avs.cd_solicitacao_ordem_srv(+)
                    and so.cd_os=osec.cd_os(+)
                    and so.cd_oficina=o.cd_oficina
                    and so.cd_espec=me.cd_espec(+)
                    AND so.cd_os = {os}"""
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
            <h6 style="font-weight: bold; font-family: Arial; color: #000000;">Comprovante de solicitação de serviço de Manutenção - OS: {ress[5]}</h6>
            <table class="table table-bordered border-primary">
                  <tr>    
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;">Solicitante</th>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" >Ramal</th>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;">Data</th>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;">Situação</th>
                  </tr>
                <tr>
                 
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial" >{ress[1]}</td>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial; text-transform: uppercase;" >{ress[6]}</td>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial">{ress[0]}</td>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial; text-transform: uppercase;">{ress[11]}</td>
                </tr>
                <tr>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;">Plaqueta - Bem patrimônial</th>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="3">Tipo de OS</th>
                  </tr>
                <tr>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial">{ress[10]}</td>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial" colspan="3">{ress[7]}</td>                
                </tr>
                <tr>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="2">Setor</th>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="2">Oficina</th>
                    
                    
                  </tr>
                 
                <tr>
                 <td scope="col" style="color: #000000; font-size: 11; font-family: Arial" colspan="2">{ress[8]}</td>
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial" colspan="2">{ress[18]}</td>
                  
                 
                 
                </tr>
                <tr>
             
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="4">Localidade</th>
                   
                    
                  </tr>
                 
                <tr>
                 
                  <td scope="col" style="color: #000000; font-size: 11; font-family: Arial" colspan="4" >{ress[9]}</td>
                  
                 
                </tr>
                <tr>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="4">Serviço  Solicitado</th>
                  </tr>
                 
              
                <tr>
                  <td scope="col" style="word-break: break-all; color: #000000; font-size: 11; font-family: Arial; text-transform: uppercase;" colspan="4">{ress[4]}</td>
                 
                </tr>
                <tr>
                    <th scope="col" style="font-weight: bold; font-size: 14; font-family: Arial; color: #000000; background-color: #c1c1c1; -webkit-print-color-adjust: exact;" colspan="4">Descrição do Serviço  Solicitado</th>
                  </tr>
              
                <tr>
                  <td scope="col" style="word-break: break-all; color: #000000; font-size: 11; font-family: Arial; text-transform: uppercase;" colspan="4">{ress[3]}</td>
                 
                </tr>
                
                
              </table>
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
            
        pdfkit.from_string(s, output_path = "C:\pdf_arquivos\imprimir.pdf")
        
        imprimir()
    
        inserir_os()   
        
    else:
        time.sleep(3)
        data_e_hora_atuais = datetime.now()
        print("Nenhuma os para ser impressa", data_e_hora_atuais, "oficina: " , cod_oficina)

# pyinstaller --onefile ger_impressao.py


    






    
