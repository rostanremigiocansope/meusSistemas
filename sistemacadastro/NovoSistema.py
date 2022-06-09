import tkinter as tk
from tkinter  import ttk
import tkinter.messagebox
import pymysql
from tkinter import END
#==Construção de Calendário==>
import datetime
from tkcalendar import*
import babel.numbers
#============================>
import time
import datetime
#Criar Planilha==============>
#Criar Planilha==============>
import xlsxwriter as XW
import mysql.connector as mq
#====Enviar Imail======================>
from tkinter import *
from tkinter import filedialog
import smtplib
from email.message import EmailMessage
#======================================>
attachments = []
#======================================>
from PIL import Image, ImageTk
#=========Criar documento====
from docx import Document
#============================

class PrimeiraPagina(tk.Frame):
    def __init__(self,parent,controller):
        tk.Frame.__init__(self,parent)

        #================================
        def senha():
            if entrySenha.get()=="batalhaocuamba":
                controller.show_frame(SegundaPagina)
                limp()
            else:
                tkinter.messagebox.showinfo("ERRO","Introduza Senha Correcta")

        def limp():
            entrySenha.delete(0, "end")
            
        #================================
        framePrincipal = tk.Frame(self, width=1350, height=700, bd=3, relief="raise")
        framePrincipal.grid()
        #=======SubFrames========
        frameTitulo1=tk.Frame(framePrincipal, bd=2, width=1340, height=150, padx=2)
        frameTitulo1.grid()

        lbFrameFarda=tk.LabelFrame(frameTitulo1, text="Exercito", width=205, height=140)
        lbFrameFarda.place(x=0,y=0)
        farda = Image.open("imagens/farda2.jpg")
        photo = ImageTk.PhotoImage(farda)
        label = tk.Label(lbFrameFarda, image=photo)
        label.image=photo
        label.place(x=10,y=0)


        lbFrametitulo=tk.LabelFrame(frameTitulo1, text="Unidade", width=780, height=140)
        lbFrametitulo.place(x=210,y=0)

        labelTitulo2=tk.Label(lbFrametitulo, font=("Times new Roman", 38,"bold"),fg="gray", text="Batalhão de Infantaria de Cuamba")
        labelTitulo2.place(x=0, y=0)
        #=============
        lbFrameSimbolo=tk.LabelFrame(frameTitulo1, text="Moçambique", width=330, height=140)
        lbFrameSimbolo.place(x=1000,y=0)
    
        bandeira = Image.open("imagens/bandeira3.png")
        photo = ImageTk.PhotoImage(bandeira)
        label = tk.Label(lbFrameSimbolo, image=photo)
        label.image=photo
        label.place(x=10,y=0)

        emblema = Image.open("imagens/emblema3.png")
        photo = ImageTk.PhotoImage(emblema)
        label = tk.Label(lbFrameSimbolo, image=photo)
        label.image=photo
        label.place(x=190,y=0)
        #==============

        frameTitulo2=tk.LabelFrame(framePrincipal, bd=2, width=1330, height=70, padx=2)
        frameTitulo2.grid()
        
        labelTitulo=tk.Label(frameTitulo2, font=("Times new Roman", 30,"bold"),fg="gray", text="Mapa geral de dados biográfico do Batalhão de Infantaria de Cuamba")
        labelTitulo.place(x=10, y=0)

        frameTitulo3=tk.LabelFrame(framePrincipal, bd=2, width=1330, height=450, padx=2)
        frameTitulo3.grid()


        lbFrameArquivo=tk.LabelFrame(frameTitulo3, text="Arquivo Digital", width=980, height=430)
        lbFrameArquivo.place(x=10,y=0)
        
        arquivo = Image.open("imagens/arquivo2.png")
        photo = ImageTk.PhotoImage(arquivo)
        label = tk.Label(lbFrameArquivo, image=photo)
        label.image=photo
        label.place(x=0,y=0)
        #================================
        lblFmSenha=tk.LabelFrame(frameTitulo3, width=400, height=200)
        lblFmSenha.place(x=400,y=150)

        labelSenha=tk.Label(lblFmSenha, font=("Times new Roman", 30,"bold"),fg="purple", text="Senha de Entrada")
        labelSenha.place(x=40, y=0)

        entrySenha=tk.Entry(lblFmSenha, font=("Times New Roman", 15,"bold"),show="*",bg="powder blue",width=15, bd=3)
        entrySenha.place(x=120, y=70)


        btEntrar = tk.Button(lblFmSenha,font=("Times new Roman",15,"bold"),bg="#ADD8E6",bd=3, text="Confirmar", relief="raise", command=senha)
        btEntrar.place(x=150, y=130)
        #================================

        lbFrameMoz=tk.LabelFrame(frameTitulo3, text="Moçambique", width=320, height=430)
        lbFrameMoz.place(x=1000,y=0)
        
        mapa = Image.open("imagens/mapa.jpg")
        photo = ImageTk.PhotoImage(mapa)
        label = tk.Label(lbFrameMoz, image=photo)
        label.image=photo
        label.place(x=40,y=0)
        
        #================================

        
        #================================
        #btconfirmar = tk.Button(self,font=("Times new Roman",15,"bold"), text="Seguinte",command=lambda: controller.show_frame(SegundaPagina))
        #btconfirmar.place(x=10, y=600)
        #================================
        #==============FIM===============      
class SegundaPagina (tk.Frame):
    def __init__(self,parent,controller):
        tk.Frame.__init__(self,parent)

        #Enviar imail===============================
        #=================Criacão de Planilhas, envio de email========================================>>>
        def criarplanilha():
            if edNomear.get()=="":
                tkinter.messagebox.showinfo(title="ERRO", message="Nomeie a Planilha")
            else:
                nomedeplanilha=str(nomear.get())
                wb = XW.Workbook("{}.Xlsx".format(nomedeplanilha))
                sh = wb.add_worksheet("Dados_de_BD")
                
                conn = mq.connect(host="localhost", user="root", passwd="", database="cuamba")
                cur = conn.cursor()
                query= "Select * from batalhaocuamba"
                cur.execute(query)
                res = cur.fetchall()

                head=["N/O","patente",
                "nome","cargo","nascimentodata","mesnascimento",
                "anonascimento",
                "dataincorporacao",
                "mesincorporacao",
                "anoincorporacao",
                "naturalidade",
                "habilitacoes",
                "datapromocao",
                "ordemservico",  
                "contactoindividual",
                "contactofamiliar",
                "especialidade",
                "sector",
                "situacao",
                "localizacao"]

                for k in range(len(head)):
                    sh.write(0,k,head[k])
                
                row=1

                for k in res:
                    for q in range(len(k)):
                        sh.write(row, q, k[q])
                    row=row+1
                tkinter.messagebox.showinfo("Rostan Cansope System","Planilha Criada...")
                conn.close()
                wb.close()
            
            

        def attachFile():
            filename = filedialog.askopenfilename(initialdir="c:/", title="please select a file")
            attachments.append(filename)
            notif.config(fg="blue", text ="Adicionado" + str(len(attachments)) + " Arquivos")
            
        def reset():
            entryemail.delete(0, "end")
            entrysenha.delete(0, "end")
            entryrec.delete(0, "end")
            entryassunt.delete(0, "end")
            entrymens.delete(0, "end")
            edNomear.delete(0, "end")

        def send():
            #try:
                msg      = EmailMessage()
                username = temp_username.get()
                password = temp_password.get()
                to       = temp_receiver.get()
                subject  = temp_subject.get()
                body     = temp_body.get()
                msg["subject"] = subject
                msg["from"] = username
                msg["to"] = to
                msg.set_content(body)

                for filename in attachments:
                    filetype = filename.split(".")
                    filetype = filetype[1]
                    if filetype == "jpg" or filetype =="JPG" or filetype == "png" or filetype =="PNG":
                        import imghdr
                        with open(filename, "rb") as f:
                            file_data = f.read()
                            image_type = imghdr.what(filename)
                        msg.add_attachment(file_data, maintype="image", sutype=image_type, filename=f.name)
                    else:
                        with open(filename, "rb") as f:
                            file_data = f.read()
                        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=f.name)
                
                if username=="" or password=="" or to=="" or subject=="" or body=="":
                    notif.config(text="Todos Campos Obrigatórios", fg="red")
                    return
                else:
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    server.login(username, password)
                    server.send_message(msg)
                    notif.config(text="Email enviado", fg="blue")

            #except:
                #notif.config(text="Erro ao enviar email", fg="red")
        #==============================================================================================================>
        def ExibirTv1():
            try:
                sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
                cur = sqlCon.cursor()
                cur.execute("select * from batalhaocuamba order by nome")
                result=cur.fetchall()
                if len(result) != 0:
                    tv1.delete(* tv1.get_children())
                    for row in result:
                        tv1.insert("",END, values=row)
                    sqlCon.commit()
                sqlCon.close()
            except:
                tkinter.messagebox.showinfo(title="ERRO", message="Por favor ligue o WampServer")

        def pessoasInfoTV1(ev):
            try:
                time.sleep(0.3)  
                verInfo=tv1.focus()
                lerData=tv1.item(verInfo)
                row=lerData["values"]
                
                NO2.set(row[0]),
                PATENTE.set(row[1]),
                NOME.set(row[2]),
                CARGO.set(row[3]),

                NASCIMENTOdata.set(row[4]),
                NASCIMENTOmes.set(row[5]),
                NASCIMENTOano.set(row[6]),
                INCORPORACAOdata.set(row[7]),
                INCORPORACAOmes.set(row[8]),
                INCORPORACAOano.set(row[9]),
                
                NATURALIDADE.set(row[10]),
                ABILITACOES.set(row[11]),
                DATAPROMO.set(row[12]),
                ORDEMSERVICO.set(row[13]),
                CONTACTO.set(row[14]),
                CONTACTOFAMILIAR.set(row[15]),
                ESPECIALIDADE.set(row[16]),
                SECTOR.set(row[17]),
                SITUACAO.set(row[18]),
                LOCALIZACAO.set(row[19]),

                
                Ano=int(NASCIMENTOano.get())
                A=data.year
                hoje=int(A)
                idade = hoje-Ano
                IDADE=str(idade)
                IDADEACTUAL.set(f"{IDADE} Anos")

                Incorporacao=int(INCORPORACAOano.get())
                Tempo=hoje-Incorporacao
                TEMPO=str(Tempo)
                TEMPOSERVICO.set(f"{TEMPO} Anos")

                if var.get()=="123":
                    doc=str(documento.get())
                    arquivo = open(f'{doc}.docx', 'a', encoding="utf-8")
                    arquivo.write("Nome Completo: "+row[2]+"\n")
                    arquivo.write("Patente: "+row[1]+"\n")
                    arquivo.write("Cargo: "+row[3]+"\n")
                    arquivo.write("Especialidade: "+row[16]+"\n")
                    arquivo.write("Sector de Trabalho: "+row[17]+"\n")
                    arquivo.write("\n\n")
                    arquivo.close()         
            except:
                tkinter.messagebox.showinfo(title="ERRO", message="Selecione um elemento")

        def Pesquisa():
            sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
            cur = sqlCon.cursor()
            if cur.execute("select * from batalhaocuamba where patente like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where nome like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where cargo like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where habilitacoes like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where especialidade like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where sector like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where situacao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where localizacao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where nascimentodata like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where mesnascimento like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where anonascimento like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where dataincorporacao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where mesincorporacao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where anoincorporacao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where naturalidade like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where datapromocao like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where contactoindividual like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where contactofamiliar like '%"+entrypesquisaPatente.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()
            else:
                tkinter.messagebox.showinfo("Sistema Rostan", "Não Encontramos o dado Pesquisado")

        #================================
        def cadastro():
            if entrySenha.get()=="cuambafadm":
                controller.show_frame(TerceiraPagina)
                limp()
            else:
                tkinter.messagebox.showinfo("ERRO","Introduza Senha Correcta")
        def limp():
            entrySenha.delete(0, "end")

        def limparDoc():
            entryNomeDoc.delete(0, "end")

        def limpar():
            edNomear.delete(0, "end")
            
                

        #================================
        
        framePrincipal = tk.Frame(self, width=1350, height=700, bd=3, relief="raise")
        framePrincipal.grid()
        #=======SubFrames========
        frameTitulo=tk.Frame(framePrincipal, bd=2, width=1340, height=50, padx=2)
        frameTitulo.grid()

        labelTitulo=tk.Label(frameTitulo, font=("Times new Roman", 20,"bold"),fg="gray", text="Mapa geral de dados biográfico do Batalhão de Infantaria de Cuamba")
        labelTitulo.place(x=100, y=0)
        
        labelDesev=tk.Label(frameTitulo, font=("Times new Roman", 12,"bold"),fg="gray", text="Sistema Desenvolvido Por:")
        labelDesev.place(x=1100, y=0)

        labelDesevenvolvedor=tk.Label(frameTitulo, font=("Times new Roman", 12,"bold"),fg="gray", text="Rostan Peter Remígio Cansope(Tenente)")
        labelDesevenvolvedor.place(x=1050, y=20)
        #==========================
        frameDados001=tk.Frame(framePrincipal, bd=2, width=1340, height=250, padx=2)
        frameDados001.grid()

        frameDados_01=tk.LabelFrame(frameDados001, text="Dados do Militar", width=430, height=240)
        frameDados_01.place(x=0,y=0)

        NO2 = tk.StringVar()
        PATENTE = tk.StringVar()
        NOME = tk.StringVar()
        CARGO = tk.StringVar()
        NASCIMENTOdata = tk.StringVar()
        NASCIMENTOmes = tk.StringVar()
        NASCIMENTOano = tk.StringVar()
        INCORPORACAOdata = tk.StringVar()
        INCORPORACAOmes = tk.StringVar()
        INCORPORACAOano = tk.StringVar()
        NATURALIDADE = tk.StringVar()
        ABILITACOES = tk.StringVar()
        DATAPROMO = tk.StringVar()
        ORDEMSERVICO = tk.StringVar()
        IDADEACTUAL = tk.StringVar()
        TEMPOSERVICO = tk.StringVar()
        CONTACTO = tk.StringVar()
        CONTACTOFAMILIAR = tk.StringVar()
        ESPECIALIDADE = tk.StringVar()
        SECTOR = tk.StringVar()
        SITUACAO = tk.StringVar()
        LOCALIZACAO = tk.StringVar()

        var = tk.StringVar()
        documento = tk.StringVar()
        
        entryNO2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=1,bg="black", textvariable=NO2)
        entryNO2.place(x=0, y=0)

        labelpatente2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Patente")
        labelpatente2.place(x=20, y=0)
        entrypatente2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=26, bd=3, textvariable=PATENTE)
        entrypatente2.place(x=200, y=0)

        labelnome2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Nome Completo")
        labelnome2.place(x=0, y=27)
        entrynome2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=35, bd=3,textvariable=NOME)
        entrynome2.place(x=128, y=27)

        labelcargo2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Cargo/Função")
        labelcargo2.place(x=0, y=54)
        entrycargo2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=35, bd=3, textvariable=CARGO)
        entrycargo2.place(x=128, y=54)

        labelnascimento2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Nascimento")
        labelnascimento2.place(x=0, y=81)
        entrynascimentodata2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=4, bd=3,textvariable=NASCIMENTOdata)
        entrynascimentodata2.place(x=200, y=81)

        entrynascimentomes2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=6, bd=3,textvariable=NASCIMENTOmes)
        entrynascimentomes2.place(x=250, y=81)

        entrynascimentoano2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=6, bd=3,textvariable=NASCIMENTOano)
        entrynascimentoano2.place(x=310, y=81)

        labelincorp2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Incorporação")
        labelincorp2.place(x=0, y=108)
        entryincorpdata2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=4, bd=3,textvariable=INCORPORACAOdata)
        entryincorpdata2.place(x=200, y=108)

        entryincorpdata2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=6, bd=3,textvariable=INCORPORACAOmes)
        entryincorpdata2.place(x=250, y=108)

        entryincorpdata2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=6, bd=3,textvariable=INCORPORACAOano)
        entryincorpdata2.place(x=310, y=108)

        labelnaturalidade2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Naturalidade")
        labelnaturalidade2.place(x=0, y=135)
        entrynaturalidade2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=NATURALIDADE)
        entrynaturalidade2.place(x=200, y=135)

        labelHlite2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Habilitações Literárias")
        labelHlite2.place(x=0, y=162)
        entryHlite2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=ABILITACOES)
        entryHlite2.place(x=200, y=162)

        labelProm2=tk.Label(frameDados_01, font=("Times New Roman", 12,"bold"), text="Data da última Promoção")
        labelProm2.place(x=0, y=189)
        entryProm2=tk.Entry(frameDados_01, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=DATAPROMO)
        entryProm2.place(x=200, y=189)
        #======================
        frameDados_02=tk.LabelFrame(frameDados001, text="Dados do Militar", width=430, height=240)
        frameDados_02.place(x=440,y=0)

        labelOrdemServico2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="O/Serviço")
        labelOrdemServico2.place(x=0, y=0)
        entryOrdemServico2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=40, bd=3, textvariable=ORDEMSERVICO)
        entryOrdemServico2.place(x=88, y=0)

        labelcontacto2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Contacto Individual")
        labelcontacto2.place(x=0, y=27)
        entrycontacto2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3, textvariable=CONTACTO)
        entrycontacto2.place(x=200, y=27)

        labelcontactofami2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Contacto Fámiliar")
        labelcontactofami2.place(x=0, y=54)
        entrycontactofami2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3, textvariable=CONTACTOFAMILIAR)
        entrycontactofami2.place(x=200, y=54)

        labelespecialidade2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Especialidade")
        labelespecialidade2.place(x=0, y=81)
        entryespecialidade2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=ESPECIALIDADE)
        entryespecialidade2.place(x=200, y=81)

        labelsector2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Sector de Trabalho")
        labelsector2.place(x=0, y=108)
        entrysector2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=SECTOR)
        entrysector2.place(x=200, y=108)

        labelsituacao2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Situação")
        labelsituacao2.place(x=0, y=135)
        entrysituacao2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=SITUACAO)
        entrysituacao2.place(x=200, y=135)

        labellocalizacao2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Localização")
        labellocalizacao2.place(x=0, y=162)
        entrylocalizacao2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=LOCALIZACAO)
        entrylocalizacao2.place(x=200, y=162)
        labelidade2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Idade Actual")
        labelidade2.place(x=0, y=189)
        entryidade2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=8, bd=3,textvariable=IDADEACTUAL)
        entryidade2.place(x=100, y=189)

        labeltempo2=tk.Label(frameDados_02, font=("Times New Roman", 12,"bold"), text="Tempo Serviço")
        labeltempo2.place(x=200, y=189)
        entrytempo2=tk.Entry(frameDados_02, font=("Times New Roman", 12,"bold"),width=8, bd=3,textvariable=TEMPOSERVICO)
        entrytempo2.place(x=310, y=189)
        #========
        criarDoc=tk.LabelFrame(frameDados001,text="Criar Doc", width=90, height=240)
        criarDoc.place(x=880,y=0)

        c = Checkbutton(criarDoc, text="Selecionar", variable=var, onvalue="123", offvalue="")
        c.deselect()
        c.place(x=0, y=40)

        labelNomear=tk.Label(criarDoc, font=("Times New Roman", 12,"bold"), text="Nomear")
        labelNomear.place(x=5, y=80)

        entryNomeDoc=tk.Entry(criarDoc, font=("Times New Roman", 12,"bold"),width=9, bd=3, textvariable=documento)
        entryNomeDoc.place(x=0, y=120)

        btlimparDoc=tk.Button(criarDoc, font=("Times New Roman", 8,"bold"),text="Limpar",bg="gray", width=5, bd=3, relief="raise", command=limparDoc)
        btlimparDoc.place(x=15, y=160)
        #############################

        frameCalendario=tk.LabelFrame(frameDados001, text="Calendário", width=350, height=240)
        frameCalendario.place(x=980,y=0)

        data=datetime.datetime.now()
        d=data.day
        m=data.month
        a=data.year
        cal=Calendar(frameCalendario,selectmode="day",year=a, month=m,day=d,date_pattertn="dd/mm/y", font=("arial",11,"bold"))
        cal.place(x=32, y=0)
        #============================
        nomear = tk.StringVar()

        frame_3=tk.Frame(framePrincipal, bd=2, width=1340, height=120, padx=2)
        frame_3.grid()

        framePlanilhaBD=tk.LabelFrame(frame_3, text="Planilha de Todos Dados", width=170, height=110)
        framePlanilhaBD.place(x=0,y=0)

        lblNomear=tk.Label(framePlanilhaBD, font=("Times New Roman",10,"bold"), text="Nomear")
        lblNomear.place(x=60,y=10)
        edNomear=tk.Entry(framePlanilhaBD,font=("arial",10),bg="powder blue",width=15, bd=3, textvariable=nomear)
        edNomear.place(x=25, y=30)
        btpesquisarcriarPlanilha=tk.Button(framePlanilhaBD, font=("Times New Roman", 8,"bold"),text="Ok",bg="gray", width=4, bd=3, relief="raise", command=criarplanilha)
        btpesquisarcriarPlanilha.place(x=25, y=60)

        btlimpar=tk.Button(framePlanilhaBD, font=("Times New Roman", 8,"bold"),text="Limpar",bg="gray", width=5, bd=3, relief="raise", command=limpar)
        btlimpar.place(x=90, y=60)

        temp_username = tk.StringVar()
        temp_password = tk.StringVar()
        temp_receiver = tk.StringVar()
        temp_subject = tk.StringVar()
        temp_body = tk.StringVar()

        frameEnviarEmail=tk.LabelFrame(frame_3, text="Enviar Email", width=650, height=110)
        frameEnviarEmail.place(x=180,y=0)

        labelemail=tk.Label(frameEnviarEmail, font=("Times New Roman",10,"bold"), text="E-mail")
        labelemail.place(x=0,y=0)
        entryemail=tk.Entry(frameEnviarEmail,font=("arial",10),bg="powder blue",width=23, bd=3,textvariable=temp_username)
        entryemail.place(x=70, y=0)

        labelsenha=tk.Label(frameEnviarEmail, font=("Times New Roman",10,"bold"), text="Senha")
        labelsenha.place(x=0,y=25)
        entrysenha=tk.Entry(frameEnviarEmail,font=("arial",10),show="*",bg="powder blue",width=23, bd=3,textvariable=temp_password)
        entrysenha.place(x=70, y=25)

        labelrec=tk.Label(frameEnviarEmail, font=("Times New Roman",10,"bold"), text="Receptor")
        labelrec.place(x=0,y=50)
        entryrec=tk.Entry(frameEnviarEmail,font=("arial",10),bg="powder blue",width=23, bd=3, textvariable=temp_receiver)
        entryrec.place(x=70, y=50)

        labelassunt=tk.Label(frameEnviarEmail, font=("Times New Roman",10,"bold"), text="Assunto")
        labelassunt.place(x=280,y=0)
        entryassunt=tk.Entry(frameEnviarEmail,font=("arial",10),bg="powder blue",width=23, bd=3,textvariable=temp_subject)
        entryassunt.place(x=350, y=0)

        labelmens=tk.Label(frameEnviarEmail, font=("Times New Roman",10,"bold"), text="Mensagem")
        labelmens.place(x=280,y=25)
        entrymens=tk.Entry(frameEnviarEmail,font=("arial",10),bg="powder blue",width=23, bd=3,textvariable=temp_body)
        entrymens.place(x=350, y=25)

        notif = tk.Label(frameEnviarEmail,width=20, text="",bg="gray", font=("Calibri", 11))
        notif.place(x=351, y=55)

        botaoAnexo=tk.Button(frameEnviarEmail, font=("Times New Roman", 8,"bold"),text="Anexo",bg="gray", width=10, bd=3, relief="raise", command=attachFile)
        botaoAnexo.place(x=550, y=0)

        botaoEnviar=tk.Button(frameEnviarEmail, font=("Times New Roman", 8,"bold"),text="Enviar",bg="gray", width=10, bd=3, relief="raise", command=send)
        botaoEnviar.place(x=550, y=25)

        botaopesquisarcriarPlanilha=tk.Button(frameEnviarEmail, font=("Times New Roman", 8,"bold"),text="Limpar",bg="gray", width=10, bd=3, relief="raise")
        botaopesquisarcriarPlanilha.place(x=550, y=50)
        #============

        frameLivre=tk.LabelFrame(frame_3, text="Livre", width=490, height=110)
        frameLivre.place(x=840,y=0)

        botaoactualizar=tk.Button(frameLivre, font=("Times New Roman", 8,"bold"),text="Actualizar Dados",bg="gray", bd=3, relief="raise", command=ExibirTv1)
        botaoactualizar.place(x=0, y=0)

        framePesquisarPatente=tk.LabelFrame(frameLivre, text="Pesquisar Dodos", width=200, height=50)
        framePesquisarPatente.place(x=0,y=40)
        entrypesquisaPatente=tk.Entry(framePesquisarPatente, font=("Times New Roman", 12,"bold"),width=16, bd=3)
        entrypesquisaPatente.place(x=0, y=0)
        btpesquisaPatente=tk.Button(framePesquisarPatente, font=("Times New Roman", 8,"bold"),text="Ok",bg="gray", width=4, bd=3, relief="raise",command=Pesquisa)
        btpesquisaPatente.place(x=150, y=0)

        frameSenha=tk.LabelFrame(frameLivre, text="Senha para Pagina de Cadastro", width=250, height=90)
        frameSenha.place(x=210,y=0)

        labelSenha=tk.Label(frameSenha, font=("Times new Roman", 10,"bold"),fg="gray", text="Digite a Senha")
        labelSenha.place(x=30, y=0)

        entrySenha=tk.Entry(frameSenha, font=("Times New Roman", 12,"bold"),show="*",width=10, bd=3)
        entrySenha.place(x=140, y=0)

        btVolt = tk.Button(frameSenha,font=("Times new Roman",8,"bold"), text="Voltar",bg="gray", bd=3, relief="raise",width=9, command=lambda: controller.show_frame(PrimeiraPagina))
        btVolt.place(x=20, y=40)
        
        botaoConf=tk.Button(frameSenha, font=("Times New Roman", 8,"bold"),text="Confirmar",bg="gray", bd=3,width=9, relief="raise",command=cadastro)
        botaoConf.place(x=150, y=40)

        
        frame_4=tk.Frame(framePrincipal, bd=2, width=1340, height=280, padx=2)
        frame_4.grid()

        #scroll_y=Scrollbar(frame_4, orient=VERTICAL)
        tv1=ttk.Treeview(frame_4, column=("NO","patente","nome","cargo","nascdata","nascimes","nasciano",
                                          "incordata","incormes","incorano","naturalidade","hab","datapromocao",
                                          "ordemservico","contacto","contactofami","Especialidd","sector","situacao",
                                          "localizacao"), show="headings", height=13)#, yscrollcommand=scroll_y.set)
        #scroll_y.pack(side=RIGHT, fill=Y)

        tv1.column("NO", width=40)
        tv1.column("patente", width=70)
        tv1.column("nome", width=150)
        tv1.column("cargo", width=90)
        tv1.column("nascdata",width=30)
        tv1.column("nascimes", width=30)
        tv1.column("nasciano", width=40)
        tv1.column("incordata", width=30)
        tv1.column("incormes", width=30)
        tv1.column("incorano", width=40)
        tv1.column("naturalidade", width=70)
        tv1.column("hab", width=80)
        tv1.column("datapromocao", width=50)
        tv1.column("ordemservico", width=90)
        tv1.column("contacto", width=70)
        tv1.column("contactofami", width=90)
        tv1.column("Especialidd", width=70)
        tv1.column("sector", width=90)
        tv1.column("situacao", width=70)
        tv1.column("localizacao", width=70)

        tv1.heading("NO", text="N/O")
        tv1.heading("patente", text="Patente")
        tv1.heading("nome", text="Nome Completo")
        tv1.heading("cargo", text="Função")
        tv1.heading("nascdata",text="D")
        tv1.heading("nascimes", text="M")
        tv1.heading("nasciano", text="A")
        tv1.heading("incordata", text="D")
        tv1.heading("incormes", text="M")
        tv1.heading("incorano", text="A")
        tv1.heading("naturalidade", text="Nat")
        tv1.heading("hab", text="Nível")
        tv1.heading("datapromocao", text="Promocão")
        tv1.heading("ordemservico",text="O/Serviço")
        tv1.heading("contacto", text="Contact")
        tv1.heading("contactofami", text="Fámilia")
        tv1.heading("Especialidd", text="Especialidade")
        tv1.heading("sector", text="Sector")
        tv1.heading("situacao", text="Situação")
        tv1.heading("localizacao", text="Localização")
        tv1.place(x=0, y=10)
        tv1.bind("<ButtonRelease-1>", pessoasInfoTV1)   
        #================================
        #btVoltar = tk.Button(self,font=("Times new Roman",15,"bold"), text="Voltar", command=lambda: controller.show_frame(PrimeiraPagina))
        #btVoltar.place(x=180, y=600)
        #btSeguinte = tk.Button(self,font=("Times new Roman",15,"bold"), text="Seguinte", command=lambda: controller.show_frame(TerceiraPagina))
        #btSeguinte.place(x=280, y=600)
        #================================
        #==============FIM===============
class TerceiraPagina(tk.Frame):
    def __init__(self,parent,controller):
        tk.Frame.__init__(self,parent)

        #=====Inserir Cadastro===============================================================================================>
        def inserir():
            try:
                time.sleep(0.4)
                if entrynome.get()=="":
                    tkinter.messagebox.showinfo(title="ERRO", message="Insira Nome Completo") 

                else:
                    sqlCon=pymysql.connect(host="localhost",user="root", password="",database="cuamba")
                    cur=sqlCon.cursor()
                    cur.execute("insert into batalhaocuamba values (default,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(       
                    patente.get(),
                    nome.get(),
                    cargo.get(),
                    nascimentodata.get(),
                    mesnascimento.get(),
                    anonascimento.get(),
                    dataincorporacao.get(),
                    mesincorporacao.get(),
                    anoincorporacao.get(),
                    naturalidade.get(),
                    habilitacoes.get(),
                    datapromocao.get(),
                    ordemservico.get(),  
                    contactoindividual.get(),
                    contactofamiliar.get(),
                    especialidade.get(),
                    sector.get(),
                    situacao.get(),
                    localizacao.get(),
                    ))
                    sqlCon.commit()
                    sqlCon.close()
                    limpar()
                    time.sleep(0.4)
                    tkinter.messagebox.showinfo("MySql connection", "Inserido com sucesso")
            except:
                tkinter.messagebox.showinfo(title="ERRO", message="ERRO NA INSERÇÂO!")
                
        #=====Exibir Cadastro==========
        def Exibir():
            try:
                sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
                cur = sqlCon.cursor()
                cur.execute("select * from batalhaocuamba order by nome")
                result=cur.fetchall()
                if len(result) != 0:
                    tv2.delete(* tv2.get_children())
                    for row in result:
                        tv2.insert("",END, values=row)
                    sqlCon.commit()
                sqlCon.close()
            except:
                tkinter.messagebox.showinfo(title="ERRO", message="Por favor ligue o WampServer")

        def pessoasInfoTV2(ev):
            time.sleep(0.3)
            try:
                verInfo=tv2.focus()
                lerData=tv2.item(verInfo)
                row=lerData["values"]
                NO.set(row[0]),
                patente.set(row[1]),
                nome.set(row[2]),
                cargo.set(row[3]),
                nascimentodata.set(row[4]),
                mesnascimento.set(row[5]),
                anonascimento.set(row[6]),
                dataincorporacao.set(row[7]),
                mesincorporacao.set(row[8]),
                anoincorporacao.set(row[9]),
                naturalidade.set(row[10]),
                habilitacoes.set(row[11]),
                datapromocao.set(row[12]),
                ordemservico.set(row[13]),  
                contactoindividual.set(row[14]),
                contactofamiliar.set(row[15]),
                especialidade.set(row[16]),
                sector.set(row[17]),
                situacao.set(row[18]),
                localizacao.set(row[19]),
            except:
                tkinter.messagebox.showinfo(title="ERRO", message="Selecione um elemento")


        def Pesquisa2():
            sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
            cur = sqlCon.cursor()
            if cur.execute("select * from batalhaocuamba where patente like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where nome like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where cargo like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv1.delete(*tv1.get_children())
                    for row in result:
                        tv1.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where habilitacoes like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where especialidade like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where sector like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where situacao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where localizacao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()


            elif cur.execute("select * from batalhaocuamba where nascimentodata like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2 .get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where mesnascimento like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where anonascimento like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where dataincorporacao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where mesincorporacao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where anoincorporacao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where naturalidade like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where datapromocao like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where contactoindividual like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2 .insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()

            elif cur.execute("select * from batalhaocuamba where contactofamiliar like '%"+entrypesquisaPatente2.get()+"%'"):
                result = cur.fetchall()
                if len(result) != 0:
                    tv2.delete(*tv2.get_children())
                    for row in result:
                        tv2.insert("", END, values=row)
                    sqlCon.commit()
                sqlCon.close()
            else:
                tkinter.messagebox.showinfo("Sistema Rostan", "Não Encontramos o dado Pesquisado")
        #=====Alterar Cadastro===========================================================================================>
        def Alterar():
            sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
            cur = sqlCon.cursor()
            cur.execute("update batalhaocuamba set patente=%s,nome=%s,cargo=%s,nascimentodata=%s,mesnascimento=%s,anonascimento=%s,dataincorporacao=%s,mesincorporacao=%s,anoincorporacao=%s,naturalidade=%s,habilitacoes=%s,datapromocao=%s, ordemservico=%s,contactoindividual=%s,contactofamiliar=%s,especialidade=%s,sector=%s,situacao=%s, localizacao=%s where NO=%s", (

                patente.get(),
                nome.get(),
                cargo.get(),
                nascimentodata.get(),
                mesnascimento.get(),
                anonascimento.get(),
                dataincorporacao.get(),
                mesincorporacao.get(),
                anoincorporacao.get(),
                naturalidade.get(),
                habilitacoes.get(),
                datapromocao.get(),
                ordemservico.get(),  
                contactoindividual.get(),
                contactofamiliar.get(),
                especialidade.get(),
                sector.get(),
                situacao.get(),
                localizacao.get(),
                NO.get(),
            ))
            sqlCon.commit()
            sqlCon.close()
            tkinter.messagebox.showinfo("Data Entry Form", "Gravação actualizada com sucesso")
        #=====Excluir Cadastro==========
        def Excluir():
            sqlCon = pymysql.connect(host="localhost", user="root", password="", database="cuamba")
            cur = sqlCon.cursor()
            cur.execute("delete from batalhaocuamba where NO=%s", NO.get())
            sqlCon.commit()
            sqlCon.close()
            tkinter.messagebox.showinfo("Data Entry Form", "Dado Excluido com sucesso")
        #=====Limpar Cadastro==========
        def limpar():
            entryNO.delete(0, "end")
            patente.set("")
            entrynome.delete(0, "end")
            entrycargo.delete(0, "end")
            nascimentoData.set("")
            nascimentoMes.set("")
            nascimentoAno.set("")
            incorporacaoData.set("")
            incorporacaoMes.set("")
            incorporacaoAno.set("")
            entrynaturalidade.delete(0, "end")
            Hliterarias.set("")
            entryProm.delete(0, "end")
            entryOrdemServico.delete(0, "end")
            entrycontacto.delete(0, "end")
            entrycontactofami.delete(0, "end")
            entryespecialidade.delete(0, "end")
            entrysector.delete(0, "end")
            situacao.set("")
            entrylocalizacao.delete(0, "end")
    
        lbltitulo = tk.Label(self,font=("Times new Roman",20,"bold"), text="Terceira Pagina")
        lbltitulo.place(x=180, y=10)  
        #================================
        frame_pagina_01_2=tk.Frame(self, width=1350, height=700, bd=3, relief="raise")
        frame_pagina_01_2.grid()

        frame_1_2=tk.Frame(frame_pagina_01_2, bd=2, width=1340, height=50, padx=2)
        frame_1_2.grid()

        labelInfo=tk.Label(frame_1_2, font=("Times new Roman", 20,"bold"),fg="gray", text="Mapa geral de dados biográfico do Batalhão de Infantaria de Cuamba/Campo de Cadastro")
        labelInfo.place(x=120, y=0)
        #==============>>>
        frame_2_2=tk.Frame(frame_pagina_01_2, bd=2, width=1340, height=250, padx=2)
        frame_2_2.grid()
        #==============>>>
        nasc = tk.StringVar()

        NO = tk.StringVar()
        patent = tk.StringVar()
        nome = tk.StringVar()
        cargo = tk.StringVar()


        nascimentodata = tk.StringVar()
        mesnascimento = tk.StringVar()
        anonascimento = tk.StringVar()

        dataincorporacao = tk.StringVar()
        mesincorporacao = tk.StringVar()
        anoincorporacao = tk.StringVar()

        naturalidade = tk.StringVar()
        habilitacoes = tk.StringVar()
        datapromocao = tk.StringVar()
        ordemservico = tk.StringVar()
        contactoindividual = tk.StringVar()
        contactofamiliar = tk.StringVar()
        especialidade = tk.StringVar()
        situacao = tk.StringVar()
        sector = tk.StringVar()
        localizacao = tk.StringVar()

        frameDados_01_2=tk.LabelFrame(frame_2_2, text="Dados do Militar", width=430, height=240)
        frameDados_01_2.place(x=0,y=0)

        entryNO=tk.Entry(frameDados_01_2, font=("Times New Roman", 12,"bold"),width=1,bg="black", textvariable=NO)
        entryNO.place(x=0, y=0)

        labelpatente=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Patente")
        labelpatente.place(x=20, y=0)
        patente=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=24, textvariable=patent)
        patente["value"]=("","Soldado","Segundo Cabo","Primeiro Cabo","Furriêl","Terceiro Sargento","Segundo Sargento","Primeiro Sargento",
                        "Subintendente","Intendente","Alfêres","Tenente","Capitão","Major","Tenente Coronel","Coronel")
        patente.current(0)
        patente.place(x=200, y=0)

        labelnome=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Nome Completo")
        labelnome.place(x=0, y=27)
        entrynome=tk.Entry(frameDados_01_2, font=("Times New Roman", 12,"bold"),width=35, bd=3, textvariable=nome)
        entrynome.place(x=128, y=27)

        labelcargo=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Cargo/Função")
        labelcargo.place(x=0, y=54)
        entrycargo=tk.Entry(frameDados_01_2, font=("Times New Roman", 12,"bold"),width=35, bd=3, textvariable=cargo)
        entrycargo.place(x=128, y=54)

        labelnascimento=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Nascimento")
        labelnascimento.place(x=0, y=81)

        nascimentoData=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=4, textvariable=nascimentodata)
        nascimentoData["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12",
                             "13","14","15","16","18","19","20","21","22","23","24","25","26",
                             "27","28","29","30","31")
        nascimentoData.current(0)
        nascimentoData.place(x=200, y=81)

        nascimentoMes=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=4, textvariable=mesnascimento)
        nascimentoMes["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12")
        nascimentoMes.current(0)
        nascimentoMes.place(x=260, y=81)

        nascimentoAno=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=6, textvariable=anonascimento)
        nascimentoAno["value"]=("","2021","2020","2019","2018","2017","2016","2015",
                                "2014","2013","2012","2011","2010","2009","2008","2007","2006","2005"
                                ,"2004","2003","2002","2001","2000","1999","1998","1997","1996","1995"
                                ,"1994","1993","1992","1991","1990","1989","1988","1987","1986","1985"
                                ,"1984","1983","1982","1981","1980","1979","1978","1977","1976","1977"
                                ,"1976","1975","1974","1973","1972","1971","1970","1969","1968","1967"
                                ,"1966","1965","1964","1963","1962","1961","1960","1959","1958","1957"
                                ,"1958","1957","1956","1955","1954","1953","1952","1951","1950","1949"
                                ,"1948","1947","1946","1945","1944","1943","1942","1941","1940","1939"
                                ,"1938","1937","1936","1935","1934","1933","1932","1931","1930")
        nascimentoAno.current(0)
        nascimentoAno.place(x=320, y=81)


        labelincorp=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Incorporação")
        labelincorp.place(x=0, y=108)

        incorporacaoData=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=4 ,textvariable=dataincorporacao)
        incorporacaoData["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12",
                             "13","14","15","16","18","19","20","21","22","23","24","25","26",
                             "27","28","29","30","31")
        incorporacaoData.current(0)
        incorporacaoData.place(x=200, y=108)

        incorporacaoMes=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=4, textvariable=mesincorporacao)
        incorporacaoMes["value"]=("","01","02","03","04","05","06","07","08","09","10","11","12")
        incorporacaoMes.current(0)
        incorporacaoMes.place(x=260, y=108)

        incorporacaoAno=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=6, textvariable=anoincorporacao)
        incorporacaoAno["value"]=("","2021","2020","2019","2018","2017","2016","2015",
                                "2014","2013","2012","2011","2010","2009","2008","2007","2006","2005"
                                ,"2004","2003","2002","2001","2000","1999","1998","1997","1996","1995"
                                ,"1994","1993","1992","1991","1990","1989","1988","1987","1986","1985"
                                ,"1984","1983","1982","1981","1980","1979","1978","1977","1976","1977"
                                ,"1976","1975","1974","1973","1972","1971","1970","1969","1968","1967"
                                ,"1966","1965","1964","1963","1962","1961","1960","1959","1958","1957"
                                ,"1958","1957","1956","1955","1954","1953","1952","1951","1950","1949"
                                ,"1948","1947","1946","1945","1944","1943","1942","1941","1940","1939"
                                ,"1938","1937","1936","1935","1934","1933","1932","1931","1930")
        incorporacaoAno.current(0)
        incorporacaoAno.place(x=320, y=108)

        labelnaturalidade=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Naturalidade(Prov-Dist)")
        labelnaturalidade.place(x=0, y=135)
        entrynaturalidade=tk.Entry(frameDados_01_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=naturalidade)
        entrynaturalidade.place(x=200, y=135)

        labelHlite=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Habilitações Literárias")
        labelHlite.place(x=0, y=162)

        Hliterarias=ttk.Combobox(frameDados_01_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=24, textvariable=habilitacoes)
        Hliterarias["value"]=("","1ª Classe","2ª Classe","3ª Classe","4ª Classe","5ª Classe",
                        "6ª Classe","7ª Classe","8ª Classe","9ª Classe","10ª Classe","11ª Classe",
                        "12ª Classe","Téc Médio Prof","Licenciado","Mestrado","PHD")
        Hliterarias.current(0)
        Hliterarias.place(x=200, y=162)

        labelProm=tk.Label(frameDados_01_2, font=("Times New Roman", 12,"bold"), text="Data da última Promoção")
        labelProm.place(x=0, y=189)
        entryProm=tk.Entry(frameDados_01_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=datapromocao)
        entryProm.place(x=200, y=189)
        #===============>>>
        frameDados_02_2=tk.LabelFrame(frame_2_2, text="Dados do Militar", width=430, height=240)
        frameDados_02_2.place(x=440,y=0)

        labelOrdemServico=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="O/Serviço")
        labelOrdemServico.place(x=0, y=0)
        entryOrdemServico=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=40, bd=3,textvariable=ordemservico)
        entryOrdemServico.place(x=88, y=0)

        labelcontacto=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Contacto Individual")
        labelcontacto.place(x=0, y=27)
        entrycontacto=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=contactoindividual)
        entrycontacto.place(x=200, y=27)

        labelcontactofami=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Contacto Fámiliar")
        labelcontactofami.place(x=0, y=54)
        entrycontactofami=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=contactofamiliar)
        entrycontactofami.place(x=200, y=54)

        labelespecialidade=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Especialidade")
        labelespecialidade.place(x=0, y=81)
        entryespecialidade=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=especialidade)
        entryespecialidade.place(x=200, y=81)

        labelsector=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Sector de Trabalho")
        labelsector.place(x=0, y=108)
        entrysector=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=sector)
        entrysector.place(x=200, y=108)

        labelsituacao=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Situação")
        labelsituacao.place(x=0, y=135)

        situacao=ttk.Combobox(frameDados_02_2, state="readonly",
                                           font=("Times New Roman",12,"bold"), width=24, textvariable=situacao)
        situacao["value"]=("","Missão de Seviço","Cursos diversos","Férias","Dispensado","Ausência ilegal","Bolseiro interno",
                           "Bolseiro externo","Doentes crônicos com J. Médica",
                           "Doentes crônicos sem J. Médica","Doente internado",
                        "Doente ambulatório","Doente junto a família","Desaparecido em combate",
                           "Ferido em Combate","Desertor apartir do teatro","Desertor apartir da unidade",
                           "Paradeiro desconecido","Prisioneiro","Óbito")
        situacao.current(0)
        situacao.place(x=200, y=135)

        labellocalizacao=tk.Label(frameDados_02_2, font=("Times New Roman", 12,"bold"), text="Localização")
        labellocalizacao.place(x=0, y=162)
        entrylocalizacao=tk.Entry(frameDados_02_2, font=("Times New Roman", 12,"bold"),width=26, bd=3,textvariable=localizacao)
        entrylocalizacao.place(x=200, y=162)
        #==============>>>
        frameBotoes_2=tk.LabelFrame(frame_2_2, text="Planilha", width=90, height=240)
        frameBotoes_2.place(x=880,y=0)

        botaoInserir=tk.Button(frameBotoes_2, font=("Times New Roman", 8,"bold"),text="Inserir",bg="gray", width=10, bd=3, relief="raise", command=inserir)
        botaoInserir.place(x=0, y=30)

        botaoExibir=tk.Button(frameBotoes_2, font=("Times New Roman", 8,"bold"),text="Exibir",bg="gray", width=10, bd=3, relief="raise", command=Exibir)
        botaoExibir.place(x=0, y=55)

        botaoAlterar=tk.Button(frameBotoes_2, font=("Times New Roman", 8,"bold"),text="Alterar",bg="gray", width=10, bd=3, relief="raise", command=Alterar)
        botaoAlterar.place(x=0, y=80)

        botaoExcluir=tk.Button(frameBotoes_2, font=("Times New Roman", 8,"bold"),text="Excluir",bg="gray", width=10, bd=3, relief="raise",command=Excluir)
        botaoExcluir.place(x=0, y=105)

        botaoLimpar=tk.Button(frameBotoes_2, font=("Times New Roman", 8,"bold"),text="Limpar",bg="gray", width=10, bd=3, relief="raise", command=limpar)
        botaoLimpar.place(x=0, y=130)

        #===============>>>
        framevazio=tk.LabelFrame(frame_2_2, text="Vazio", width=340, height=240)
        framevazio.place(x=980,y=0)

        PripagVoltar = tk.LabelFrame(framevazio, width=320, height=50)
        PripagVoltar.place(x=3,y=3)

        btVolt = tk.Button(PripagVoltar,font=("Times new Roman",8,"bold"), text="Voltar",bg="gray", bd=3, relief="raise",width=12, command=lambda: controller.show_frame(SegundaPagina))
        btVolt.place(x=20, y=10)

        btVolt = tk.Button(PripagVoltar,font=("Times new Roman",8,"bold"), text="Primeira Pagina",bg="gray", bd=3, relief="raise",width=12, command=lambda: controller.show_frame(PrimeiraPagina))
        btVolt.place(x=170, y=10)
        
        #===============>>>
        frame_3_2=tk.Frame(frame_pagina_01_2, bd=2, width=1340, height=360, padx=2)
        frame_3_2.grid()

        frameTv2=tk.LabelFrame(frame_3_2, text="Dados", width=1320, height=320)
        frameTv2.place(x=0,y=0)

        framePesquisarPatente2=tk.LabelFrame(frame_3_2, text="Pesquisar Dodos", width=200, height=50)
        framePesquisarPatente2.place(x=0,y=20)
        entrypesquisaPatente2=tk.Entry(framePesquisarPatente2, font=("Times New Roman", 12,"bold"),width=16, bd=3)
        entrypesquisaPatente2.place(x=0, y=0)
        btpesquisaPatente2=tk.Button(framePesquisarPatente2, font=("Times New Roman", 8,"bold"),text="Ok",bg="gray", width=4, bd=3, relief="raise",command=Pesquisa2)
        btpesquisaPatente2.place(x=150, y=0)

        frameInfo2dados=tk.LabelFrame(frame_3_2, width=1095, height=50)
        frameInfo2dados.place(x=210,y=20)
        labelInfo2dados=tk.Label(frameInfo2dados, font=("Times new Roman", 20,"bold"),fg="gray", text="Dados biográfico do Batalhão de Infantaria de Cuamba")
        labelInfo2dados.place(x=140, y=0)

        tv2=ttk.Treeview(frame_3_2, column=("NO","patente","nome","cargo","nascdata","nascimes","nasciano",
                                          "incordata","incormes","incorano","naturalidade","hab","datapromocao",
                                          "ordemservico","contacto","contactofami","Especialidd","sector","situacao",
                                          "localizacao"), show="headings")

        tv2.column("NO", width=40)
        tv2.column("patente", width=70)
        tv2.column("nome", width=150)
        tv2.column("cargo", width=90)
        tv2.column("nascdata",width=30)
        tv2.column("nascimes", width=30)
        tv2.column("nasciano", width=40)
        tv2.column("incordata", width=30)
        tv2.column("incormes", width=30)
        tv2.column("incorano", width=40)
        tv2.column("naturalidade", width=70)
        tv2.column("hab", width=80)
        tv2.column("datapromocao", width=50)
        tv2.column("ordemservico", width=90)
        tv2.column("contacto", width=70)
        tv2.column("contactofami", width=90)
        tv2.column("Especialidd", width=70)
        tv2.column("sector", width=90)
        tv2.column("situacao", width=70)
        tv2.column("localizacao", width=70)

        tv2.heading("NO", text="N/O")
        tv2.heading("patente", text="Patente")
        tv2.heading("nome", text="Nome Completo")
        tv2.heading("cargo", text="Função")
        tv2.heading("nascdata",text="D")
        tv2.heading("nascimes", text="M")
        tv2.heading("nasciano", text="A")
        tv2.heading("incordata", text="D")
        tv2.heading("incormes", text="M")
        tv2.heading("incorano", text="A")
        tv2.heading("naturalidade", text="Nat")
        tv2.heading("hab", text="Nível")
        tv2.heading("datapromocao", text="Promocão")
        tv2.heading("ordemservico",text="O/Serviço")
        tv2.heading("contacto", text="Contact")
        tv2.heading("contactofami", text="Fámilia")
        tv2.heading("Especialidd", text="Especialidade")
        tv2.heading("sector", text="Sector")
        tv2.heading("situacao", text="Situação")
        tv2.heading("localizacao", text="Localização")
        tv2.place(x=0, y=90)
        Exibir()
        tv2.bind("<ButtonRelease-1>", pessoasInfoTV2)

        
        
        #================================
        #btVoltar = tk.Button(self,font=("Times new Roman",15,"bold"), text="Voltar", command=lambda: controller.show_frame(SegundaPagina))
        #btVoltar.place(x=180, y=600)
        #btSeguinte = tk.Button(self,font=("Times new Roman",15,"bold"), text="Home", command=lambda: controller.show_frame(PrimeiraPagina))
        #btSeguinte.place(x=280, y=600)
        #================================
        #================FIM=============
class Application(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        #criacao de window
        window = tk.Frame(self)
        window.pack()


        window.grid_rowconfigure(0, minsize =700)
        window.grid_columnconfigure(0, minsize = 1350)

        self.frames = {}
        for F in (PrimeiraPagina,SegundaPagina,TerceiraPagina):
            frame = F(window,self)
            self.frames[F] = frame
            frame.grid(row = 0, column = 0, sticky ="nsew")
            
        self.show_frame(PrimeiraPagina)

    def show_frame(self,page):
        frame = self.frames[page]
        frame.tkraise()
        
app = Application()
app.maxsize(1350,700)
app.mainloop()
