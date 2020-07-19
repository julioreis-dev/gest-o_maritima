import win32com.client
from openpyxl import load_workbook
import Planguia_funcoes


class EnviarEmail:
    def __init__(self, arquivo, aba, mes_ano, revisao):
        self.arquivo = arquivo
        self.aba = aba
        self.mes = mes_ano
        self.revisao = revisao

    def preparar_email(self):
        dados_rev = self.aba.cell(row=1, column=7).value
        mes_medicao = self.aba.cell(row=1, column=2).value
        self.aba.cell(row=1, column=16).value = ''
        dest_filename = mes_medicao + ' - Planilha Guia_R' + str(dados_rev) + '.xlsx'
        self.arquivo.save(filename=dest_filename)
        self.arquivo.close()
        return dest_filename

    def listar_revisao(self):
        ws2 = self.arquivo['Revisão']
        lista = []
        numero_linhas = self.aba.max_row
        for n in range(2, numero_linhas + 1):
            status = ws2.cell(row=n, column=3).value
            if status == 'ok':
                alter = ws2.cell(row=n, column=2).value
                lista.append(alter)
        sentenca = ''
        for w in range(0, len(lista)):
            valor = lista[w]
            sentenca = sentenca + str(valor) + '\n'
        return sentenca

    def listar_destinatario(self):
        aba1 = self.arquivo['Chaves']
        contar_dest = aba1.max_row
        lista_destinatario = []
        lista_destinatario1 = []
        lista_resultado = []
        for n in range(2, contar_dest + 1):
            status = aba1.cell(row=n, column=6).value
            if status == 'Sim':
                endereco = aba1.cell(row=n, column=1).value
                lista_destinatario.append(endereco)
            elif status == 'Copy':
                endereco2 = aba1.cell(row=n, column=1).value
                lista_destinatario1.append(endereco2)
        saudacao = Planguia_funcoes.saudar()
        destinatarios = '; '.join(lista_destinatario)
        destinatarios2 = '; '.join(lista_destinatario1)
        lista_resultado.append(destinatarios)
        corpo_email = saudacao + '\n\nSegue anexo a planilha guia de medição (Revisão ' + str(
            self.revisao) + ').\nQualquer informação a acrescentar em alguma embarcação da frota, que não constar na planilha' \
                            ', favor avisar para ajuste.\n\nATENÇÃO: PARA EMBARCAÇÕES QUE NÃO FORAM LIBERADAS ' \
                            'PARA MEDIÇÃO VER NA COLUNA (N - "Embarcação Liberada") DA ABA (MEDIÇÃO).TODAS AS ALTERAÇÕES REALIZADAS ' \
                            'NA PLANILHA GUIA SÃO REGISTRADAS NA ABA "Revisão".\n\nAs embarcações "Não Liberadas" estão sendo analisadas ' \
                            'pela equipe de coordenação e controle das informações referente a frota sob a gestão do CMAR, e em breve esses ' \
                            'equipamentos poderão sofrer alterações quanto ao seu status. Caso isso ocorra, uma nova r' \
                            'evisão será emitida com a finalidade de ajustar a planilha guia.'
        lista_resultado.append(corpo_email)
        lista_resultado.append(destinatarios2)
        return lista_resultado

    def encaminhar_email(self):
        try:
            anexo = r'C:/Users/ay4m/Desktop/Python/projetos/' + self.preparar_email()
            lista_dados = self.listar_destinatario()
            o = win32com.client.Dispatch("Outlook.Application")
            # lista_dados = contar.listar_destinatario_Petro(caminho, aba, rev)
            msg = o.CreateItem(0)
            msg.To = lista_dados[0]
            msg.CC = lista_dados[2]
            msg.BCC = ''
            msg.Subject = 'CMAR - Planilha Guia de Medição - ' + self.mes[0] + ' - REVISÃO ' + str(self.revisao)
            msg.Body = lista_dados[1] + '\nHistórico de ajustes realizados (Revisão ' + str(
                self.revisao) + '):\n' + self.listar_revisao() + '\nAtenciosamente,' \
                                                                '\nEquipe de Gerenciamento Marítimo\nLOEP/LOFF/GCI/CMAR\n' + self.mes[1]
            msg.Attachments.Add(anexo)
            msg.Send()
            return '\nPrezado usuário, Email enviado com sucesso!'
        except:
            return 'Erro ao enviar email'


def iniciar_3():
    print('Preparando email para envio....................')
    t = Planguia_funcoes.openr('Projeto_planilha_Guia_Medição.xlsx', 'Medição')
    numero_mes = t[1].cell(row=1, column=16).value
    mes = Planguia_funcoes.definir_mes(numero_mes)
    rev = t[1].cell(row=1, column=7).value
    email = EnviarEmail(t[0], t[1], mes, rev)
    resposta = email.encaminhar_email()
    print(resposta)
