import string
import gspread
import csv
import time
from oauth2client.service_account import ServiceAccountCredentials
from requests import session
from bs4 import BeautifulSoup as bs
import urllib
import os, sys, itertools, re
import datetime
import getpass
import logging

# for a in range(len(sys.argv)):
#     print(sys.argv[a])

if len(sys.argv) <= 2:
    print('sem argumentos')
    print('utilize argumentos da seguinte forma:')
    print('')
    print('python UpdateSheet.py CURSO TURMA')
    print('')
    print('Opções de curso:')
    print('- Java VIII ou Java IX')
    print('- MBA III ou MBA IV')
    sys.exit()


KEY_JAVA_VIII = '1qvlHVHBPZEAtquBi8Gh1Wnc82-C-kAButZ8bmqihT-k'
KEY_JAVA_IX = '1Fkvo3rhYFfbuL712rHyyaA3XWhkkX4_YctDSK4_BIlc'
KEY_MBA_III = '1mGRcIdVWR6OtppHgxCxFRrdGWH0Kn9IUN86wpHP_WGQ'
KEY_MBA_IV = '15DxfzavHF6HRZzmA8Dcg0CiMnB6IvBZKWmfP5u7rJow'
#PLANILHA PADRÃO
KEYs = KEY_JAVA_IX

try:
    print(sys.argv[0], sys.argv[1].upper(),sys.argv[2].upper())
    curso = sys.argv[1]
    turma = sys.argv[2]

    if 'java'.upper() == curso.upper():
        if turma.upper() in 'ix'.upper():
            KEYs = KEY_JAVA_IX
        else:
            if turma.upper() == 'viii'.upper():
                KEYs = KEY_JAVA_VIII
            else:
                print('argumento invalido')
                print('turmas possíveis para Java: IX e VIII')
                sys.exit()
    else:
        if 'mba'.upper() == curso.upper():
            if turma.upper() == 'iii'.upper():
                KEYs = KEY_MBA_III
            else:
                if turma.upper() == 'iv'.upper():
                    KEYs = KEY_MBA_IV
                else:
                    print('argumento invalido')
                    print('turmas possíveis para MBA: III e IV')
                    sys.exit()
        else:
            print('Argumentos inválidos')
            print('cursos possíveis: Java e MBA')

except:
    print('Argumentos inválidos')



USERNAME = ''
PASSWORD = ''
baseurl = "https://moodle.utfpr.edu.br/"
LOGUE = []
#CONFIGS
LOG_FILENAME = "logfile.log"
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(filename=LOG_FILENAME,level=logging.DEBUG)
sections = itertools.count()
files = itertools.count()
scope = ['https://spreadsheets.google.com/feeds']
# KEYs = '1Fkvo3rhYFfbuL712rHyyaA3XWhkkX4_YctDSK4_BIlc'
credentials = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open_by_key(KEYs)
worksheet = wks.get_worksheet(0)

# os.system("cls" if os.name == "nt" else "clear")
USERNAME = input("Nome de usuario do moodle: ")
#USUARIO = getpass.getuser("Nome de usuario do moodle: ")
# os.system("cls" if os.name == "nt" else "clear")
#SENHA = input("Senha do Moodle: ")
PASSWORD = getpass.getpass("Senha do Moodle: ")
os.system("cls" if os.name == "nt" else "clear")


class Disciplina:
    def __init__(self, codigo=None, nome_curto=None, id='0', num_atividades=0, situacao=False):
        self.codigo = codigo
        self.nome_curto = nome_curto
        self.id = id
        self.nome_completo = codigo + " - " + nome_curto
        self.num_atividades = num_atividades
        if situacao == 'FALSE':
            self.situacao = False
        else:
            self.situacao = True
        self.nome_planilha = codigo+"_"+nome_curto
        self.professor = "nome professor"
        self.email_professor = "professor@email"
        self.ATIVIDADES = []

    def get_url_curso(self):
        return baseurl + "course/view.php?id="+ str(self.id)

    def get_url_livro_notas(self):
        return baseurl + "/grade/report/grader/index.php?id=" + str(self.id)

    def __eq__(self, other):
        return self.codigo == other.codigo

    def __str__(self):
        return self.codigo + ' - '+ self.nome_curto

class Aluno:
    def __init__(self, nome=None, situacao=True, linha = 0, id='0', email=None,):
        self.nome = nome
        self.nome_planilha = self.nome
        # self.situacao = situacao
        if situacao == 'FALSE':
            self.situacao = False
        else:
            self.situacao = True
        self.linha = linha
        self.id = id
        self.email = email
        self.tutor = ''

    def set_situacao(self, situacao):
        if situacao == 'FALSE':
            self.situacao = False
        else:
            self.situacao = True

    def get_tutor(self):
        if self.tutor.find('/') == -1:
            return self.tutor
        return self.tutor[:self.tutor.find('/')]


    def __eq__(self, other):
        return self.nome == other.nome

    def __str__(self):
        return self.id + ' ' +self.nome + ' ' +self.email

class Atividade:
    def __init__(self, nome, id, tipo, link, dataitem):
        self.nome = nome
        self.id = id
        self.tipo = tipo
        self.link = link
        self.dataitem = dataitem
        self.data_entrega = ''
        self.precisa_avaliacao = ''


    def get_hiperlink(self):
        return '=HYPERLINK("'+self.link+'";"'+self.get_nome_curto()+'")'

    def get_nome_curto(self):
        return self.nome[:25]

    def __str__(self):
        return self.id + ' ' + self.nome

class Script:
    def __init__(self):
        self.session = self.login(USERNAME, PASSWORD)
        self.disciplina_coordenacao = self.obter_disciplina_coordenação()
        self.ALUNOS_MOODLE = self.obter_alunos_moodle()
        self.DISCIPLINAS = self.obter_disciplinas_ativas()
        self.linha_primeiro_aluno = 8
        self.linha_ultimo_aluno = self.linha_primeiro_aluno + len(self.ALUNOS_MOODLE) -1

    def iniciar(self):
        self.log('INICIANDO EXECUÇÃO')
        numero = self.altualizar_alunos()
        # OBTER CONF
        # PARA CADA CADA PLANILHA
        print()
        for disciplina in self.DISCIPLINAS:
            # OBTER DADOS NO MOODLE
            self.log('-----' + disciplina.nome_completo + '-------')
            self.update_curso(disciplina)
            disciplina.ATIVIDADES = self.obter_atividades_disciplina_moolde(disciplina)
            notas_alunos = self.obter_notas_alunos(disciplina)

            # csv_file = self.get_page_curso(disciplina)
            # leitor = self.get_csv(csv_file)
            # # SE PLANILHA NÃO EXISTE, CRIAR
            #planilha = self.get_planilha(disciplina)
            # # ATUALIZAR PLANILHA
            # resposta = self.obter_dados_csv(disciplina, leitor)
            # nomes = resposta[1]
            # notas_alunos = resposta[0]

            self.atualiza_planilha(disciplina, notas_alunos)
            # escreve_notas_alunos(disciplina, notas_alunos)
            print()
        self.log('FINALIZADO')

    def log(self, texto):
        texto = str(len(LOGUE)) + " - " + str(texto)
        print(texto)
        logging.info(texto)
        LOGUE.append(texto)

    def login(self, user, pwd):
        self.log('FAZENDO LOGIN NO MOODLE E CRIANDO SESSÃO')

        #adiciona headers
        headers = {
            'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Mobile Safari/537.36'
        }

        #cria sessão
        with session() as ses:
            r = ses.get(baseurl + 'login/index.php', headers=headers)

            #pesquisa os atributos faltando para login
            soup = bs(r.content, 'html5lib')
            lt = soup.find('input', attrs={'name': 'logintoken'})['value']

            #cria dados de autenticação
            authdata = {
                'anchor': '',
                'logintoken': lt,
                'username': user,
                'password': pwd
            }

            #faz o post com o headears e os dados de autenticação
            r = ses.post(baseurl + 'login/index.php', data=authdata, headers=headers)

            # verificar se o usuário conseguiu fazer login antes de devolver a sessão

            try:
                soup = bs(r.content, 'html5lib')
                username = soup.find_all('span', attrs={'class': 'usertext'})[0].contents[0]
                self.log(f'USUÁRIO LOGADO COMO {username}')
            except:
                self.log('USUÁRIO OS SENHA INVÁLIDO. TENTAR NOVAENTE')
                user = input("Nome de usuario do moodle: ")
                pwd = getpass.getpass("Senha do Moodle: ")
                ses = self.login(user, pwd)

            return ses

    def obter_disciplina_coordenação(self):
        return wks.worksheet('CONF').cell(2,3).value

    def obter_disciplinas(self):
        self.log('OBTER DISCIPLINAS NA PLANILHA CONF')
        disciplinas = []
        dados = wks.worksheet("CONF").get_all_records()
        for disc in dados:
            d = Disciplina(disc["COD"], disc["NOME"], disc["ID"], disc["ATIVIDADES"], disc["ATIVO"])
            disciplinas.append(d)
            self.log(d.__str__())
        return disciplinas

    def obter_disciplinas_ativas(self):
        self.log('OBTER DISCIPLINAS ATIVAS NA PLANILHA CONF')
        disciplinas = self.obter_disciplinas()
        disclinas_ativias = []
        for disciplina in disciplinas:
            if disciplina.situacao:
                disclinas_ativias.append(disciplina)
        return disclinas_ativias

    def obter_aluno_lista(self, nome, email) -> Aluno:
        for aluno in self.ALUNOS_MOODLE:
            if nome.upper() in aluno.nome.upper():
                # print('aluno '+ aluno.nome)
                return aluno
            if email.upper() in aluno.email.upper():
                # print('aluno ' + aluno.nome +' encontrado pelo email')
                return aluno

        self.log("aluno "+nome+" não encontrado")
        return Aluno(nome,True,0,0,email)

    def altualizar_alunos(self):
        self.log('ATUALIZANDO DADOS DA PLANILHA DOS ALUNOS')

        nomes = wks.worksheet("ALUNOS").find("Nome")
        alunos = wks.worksheet("ALUNOS").col_values(2)
        situacao = wks.worksheet("ALUNOS").col_values(1)
        emails = wks.worksheet("ALUNOS").col_values(3)
        tutor = wks.worksheet("ALUNOS").col_values(6)
        lista = []


        for a in range(len(alunos)):
            aluno = alunos[a]

        # for aluno in alunos:
        #     print('----'+aluno)
            if not (aluno == "" or aluno == "Nome" or a < nomes.row - 1):
                # print("--" + aluno)
                # al = Aluno(aluno, situação[alunos.index(aluno)], str(int(alunos.index(aluno)) + 1))
                al = self.obter_aluno_lista(aluno, emails[alunos.index(aluno)])
                al.linha = str(int(alunos.index(aluno)) + 1)
                al.set_situacao(situacao[alunos.index(aluno)])
                al.tutor = tutor[alunos.index(aluno)]
                al.nome_planilha = aluno
                # print(al.nome)
                lista.append(al)

        return lista

    def obter_alunos_ativos(self):
        alunos_ativos = []
        for aluno in self.ALUNOS_MOODLE:
            if aluno.situacao:
                alunos_ativos.append(aluno)

        return alunos_ativos

    def update_curso(self, disciplina):
        self.log('OBTENDO ID DA DISCIPLINA')
        url_search = baseurl + 'course/search.php?q=' + disciplina.codigo
        print(url_search)
        r = self.session.get(url_search)
        time.sleep(1)
        soup = bs(r.content, 'html5lib')
        try:
            nome_completo = soup.find('h3', attrs={'class': 'coursename'})
            a = nome_completo.find_all('a')[0]
            id = a['href']
            id = id[id.find('id=')+3:]
            disciplina.id = id
            self.log(f'ID DA DISCIPLINA {disciplina.codigo} LOCALIZADO: {disciplina.id}')
        except:
            self.log('ID NÃO LOCALIZADO NO MOODLE')

        self.log('OBTENDO NOME COMPLETO DA DISCIPLINA')
        try:
            nome_completo = soup.find('h3', attrs={'class': 'coursename'})
            nome_completo = nome_completo.find('a').contents[0]
            disciplina.nome_completo = nome_completo
            self.log('NOME COMPLETO: '+nome_completo)
        except:
            self.log('DISCIPLINA NÃO LOCALIZADA NO MOODLE')

        self.log('OBTENDO NOME DO PROFESSOR')
        try:
            nome_professor = soup.find('ul', attrs={'class': 'teachers'})
            nome_professor = nome_professor.find('a').contents[0]
            disciplina.professor = nome_professor
            self.log('PROFESSOR: '+nome_professor)
        except:
            self.log('PROFESSOR NÃO LOCALIZADO NO MOODLE')

    def get_page_curso(self, disciplina):
        #exportando dados do csv
        self.log('OBTENDO CSV DO CURSO')
        urlPage = 'https://moodle.utfpr.edu.br/grade/export/txt/index.php?id='+str(disciplina.id)
        urlExport = 'https://moodle.utfpr.edu.br/grade/export/txt/export.php'
        r = self.session.get(urlPage)

        soup = bs(r.content, 'html5lib')
        sesskey = soup.find('input', attrs={'name': 'sesskey'})['value']

        form = soup.find_all('form','mform')[0]
        inputs = form.find_all('input')
        export_data = {}
        for inp in inputs:
            if not inp['name'] == 'nosubmit_checkbox_controller1':
                export_data[inp['name']] = inp['value']
        export_data['display[percentage]'] = '0'
        export_data['display[letter]'] = '0'
        export_data['export_feedback'] = '0'
        export_data['decimals'] = '2'
        export_data['separator'] = 'coma'
        r = self.session.post(urlExport, data=export_data)
        return r.content

    def get_csv(self, csv_file):
        self.log('CONVERTENDO CSV EM LISTA')
        decoded_content = csv_file.decode('utf-8')
        return csv.DictReader(decoded_content.splitlines(), delimiter=',')

    def get_csv_infile(self, file):
        self.log('CONVERTENDO CSV EM DICIONÁRIO')
        with open(file, 'r') as arquivo_csv:
            return csv.DictReader(arquivo_csv, delimiter=',')

    def get_planilha(self, disciplina: Disciplina):
        self.log('OBTENDO A PLANILHA DO CURSO '+disciplina.codigo)
        cell = wks.worksheet("CONF").find(disciplina.codigo)
        recriar = wks.worksheet('CONF').cell(cell.row,6).value
        if 'TRUE' in recriar:
            try:
                self.log('PLANILHA DEVERÁ SER RECRIADA')
                wks.del_worksheet(wks.worksheet(disciplina.nome_planilha))
                self.log('PLANILHA DELETADA')
            except:
                self.log('PLANILHA NÃO ENCONTRADA PARA SER DELETADA')

        try:
            self.log('OBTENDO A PLANILHA')
            return wks.worksheet(disciplina.nome_planilha)
        except:
            self.log('A PLANILHA DEVERÁ SER CRIADA')
            return self.criar_planilha(disciplina)

    def cria_cabecalho_planilha(self, disciplina, planilha):
        # nome disciplina / professor / email / Livro de notas
        self.log('SALVANDO NA PLANILHA NOME, PROFESSOR, EMAIL E LIVRO NOTA')
        celula = '=HYPERLINK("' + disciplina.get_url_curso() + '";"' + disciplina.nome_completo + '")'
        planilha.update_cell(1, 1, celula)
        planilha.format('A1:B1', {
            "textFormat": {
                "fontSize": 24,
                "bold": True
            }
        })

        celula = disciplina.professor
        planilha.update_cell(2, 1, celula)
        celula = disciplina.email_professor
        planilha.update_cell(3, 1, celula)
        celula = '=HYPERLINK("' + disciplina.get_url_livro_notas() + '";"' + 'LIVRO DE NOTAS' + '")'
        planilha.update_cell(4, 1, celula)
        planilha.format("A2:A4", {
            "backgroundColor": {
                "red": 1.0,
                "green": 1.0,
                "blue": 1.0
            },
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {
                    "red": 255.0,
                    "green": 0.0,
                    "blue": 0.0
                },
                "fontSize": 12,
                "bold": False
            }
        })

    def grava_nomes_alunos_planilha(self, disciplina, planilha):
        self.log('GRAVANDO NA PLANILHA O NOME DOS ALUNOS')

        intervalo = self.get_intervalo(False, self.linha_primeiro_aluno, 1, len(self.ALUNOS_MOODLE))
        cell_list = planilha.range(intervalo)
        linha = 4
        for cell in cell_list:
            cell.value = '=ALUNOS!B'+str(linha)
            linha = linha + 1
        planilha.update_cells(cell_list, value_input_option='USER_ENTERED')

    def gravar_barra_atividades(self, disciplina, planilha):
        ativ = disciplina.num_atividades

        primrira_linha = 5
        inicio = 'A'+str(primrira_linha)
        fim = self.get_letra(ativ+8)+str(primrira_linha + 2)

        intervalo = inicio+':'+ fim

        # intervalo = self.get_intervalo(True, 5, 1, ativ + 7)
        # print(intervalo)

        linha1 = []
        linha2 = []
        linha3 = []

        linha1.append('NOME / ATIVIDADES:')
        linha2.append('DATAS >>')
        linha3.append('FALTA CORRIGIR>>')

        for atividade in disciplina.ATIVIDADES:
            linha1.append(atividade.get_hiperlink())

            linha2.append(atividade.data_entrega)
            linha3.append(atividade.tipo + ' - ' + str(atividade.precisa_avaliacao))

        linha1.append('ATIVIDADES')
        linha2.append('')
        linha3.append('')
        linha1.append('PROVA')
        linha2.append('')
        linha3.append('')
        linha1.append('2ª CHAMADA')
        linha2.append('')
        linha3.append('')
        linha1.append('RECUPERAÇÃO')
        linha2.append('')
        linha3.append('')
        linha1.append('NOTA FINAL')
        linha2.append('')
        linha3.append('')
        linha1.append('ANOTAR SITUAÇÃO DO ALUNO')
        linha2.append('')
        linha3.append('')
        linha1.append('DATA CONTATO')
        linha2.append('')
        linha3.append('')
        tabela = [linha1, linha2, linha3]
        # intervalo = self.get_intervalo_tabela(tabela)
        planilha.update(intervalo, tabela, value_input_option='USER_ENTERED')
        corpo =  {
            "backgroundColor": {
                "red": 0.7,
                "green": 0.7,
                "blue": 0.7
            },
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {
                    "red": 0.0,
                    "green": 0.0,
                    "blue": 0.0
                },
                "fontSize": 10,
                "bold": False
            },

        }
        planilha.format(intervalo,corpo)
        # intervalo = self.get_intervalo(True, 6, 1, ativ + 7)
        # planilha.format(intervalo, corpo)
        sheetId = wks.worksheet(disciplina.nome_planilha)._properties['sheetId']
        # print(sheetId)
        body = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheetId,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": disciplina.num_atividades + 8
                        },
                        "properties": {
                            "pixelSize": 200
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        wks.batch_update(body)

    def barrar_alunos_inativos(self, disciplina, planilha):
        self.log('BARRAR ALUNOS INATIVOS')
        # barra alunos inativos
        for aluno in self.ALUNOS_MOODLE:
            if not aluno.situacao:
                cell = planilha.find(aluno.nome_planilha)
                # intervalo = self.get_intervalo(True, cell.row, cell.col, disciplina.num_atividades + 5)
                intervalo = self.get_intervalo(True, int(aluno.linha) + 4, 1, disciplina.num_atividades + 5)
                planilha.format(intervalo, {
                    "backgroundColor": {
                        "red": 0.9,
                        "green": 0.9,
                        "blue": 0.9
                    }})
                time.sleep(1)

    def criar_planilha(self, disciplina: Disciplina):
        self.log('CRIANDO PLANILHA '+disciplina.nome_planilha)
        planilha = wks.add_worksheet(disciplina.nome_planilha, 200, 35)
        self.cria_cabecalho_planilha(disciplina, planilha)
        self.grava_nomes_alunos_planilha(disciplina, planilha)
        self.gravar_barra_atividades(disciplina, planilha)
        self.barrar_alunos_inativos(disciplina, planilha)

        return planilha

    def get_atividades(self, disciplina, leitor):
        atividades = []

        for at in leitor.fieldnames[6:]:
            if "Tarefa:" in at and not "WebConf Oficial" in at and not "Conferência" in at:
                atividades.append(at)
        return atividades

    def obter_dados_csv(self, disciplina, leitor):
        self.log('EXTRAINDO AS NOTAS DO CSV')
        alunos = {}
        nomes = []
        for coluna in leitor:
            nome = coluna['Nome'] + ' ' + coluna['Sobrenome']
            # print(nome)
            nome.encode('utf-8')
            notas = []
            for at in self.get_atividades(disciplina, leitor)[0:disciplina.num_atividades]:
                nota = coluna[at]
                notas.append(nota.replace('.', ','))

            alunos[nome] = notas
            nomes.append(nome)
        return [alunos, nomes]

    def atualiza_planilha(self, disciplina, notas_alunos):

        self.log('ATUALIZAR NOTAS NA PLANILHA')

        planilha = self.get_planilha(disciplina)

        self.gravar_barra_atividades(disciplina, planilha)
        alunos_planilha = self.altualizar_alunos()

        # for aluno in alunos_planilha:
        #     print(aluno.nome)

        tabela = []
        for aluno in alunos_planilha:
            linha = []
            try:
                linha = notas_alunos[aluno.nome]
                # print(aluno.nome, "-", notas_alunos[aluno.nome])
            except:
                for x in range(disciplina.num_atividades):
                    linha.append("")
                # print(aluno.nome + " não encontrado")
            tabela.append(linha)
            # print('linha '+str(len(linha)))


        intervalo = self.get_intervalo_tabela(tabela, 8, 2)

        # for l in tabela:
        #
        #     print(l)
        # print(tabela)

        planilha.update(intervalo, tabela, value_input_option='USER_ENTERED')
        planilha.format(intervalo, {"horizontalAlignment": "CENTER"})

        self.atividades_para_corrigir(disciplina, planilha)

        self.planilha_atualizada(disciplina)

    def planilha_atualizada(self, disciplina):
        self.log('GRAVANDO ULTIMA ATUALIZAÇÃO')
        cell = wks.worksheet('CONF').find(disciplina.codigo)
        hora = datetime.datetime.now()
        wks.worksheet('CONF').update_cell(cell.row,7,str(hora)+" - "+ USERNAME)
        self.log('ULTIMA ATUALIZAÇÃO ' + str(hora)+" - "+ USERNAME)

    def get_intervalo_tabela(self, tabela, linha_inicio, coluna_inicio):
        linhas = len(tabela)
        colunas = len(tabela[0])
        inicio = str(self.get_letra(coluna_inicio)) + str(linha_inicio)
        fim = str(self.get_letra(coluna_inicio + colunas))+str(linha_inicio+linhas)

        return inicio + ':'+fim

    def get_intervalo(self, horizontal, celInicio, colInicio, qntd):
        letras = list(string.ascii_uppercase);
        if horizontal:
            return letras[colInicio - 1]+str(celInicio)+':'+letras[colInicio+qntd-1]+str(celInicio)
        else:
            return letras[colInicio - 1] + str(celInicio) + ':' + letras[colInicio - 1] + str(celInicio+qntd)

    def get_letra(self,numero):
        n = int(numero)
        pre = ''
        for a in range(n//26):
            pre = pre+'A'
        n = n % 26
        letras = list(string.ascii_uppercase)
        return pre+letras[n-1]

    def obter_alunos_moodle(self):
        self.log('OBTENDO ALUNOS DO MOODLE')
        alunos = []
        url = 'https://moodle.utfpr.edu.br/user/index.php?id='+ self.disciplina_coordenacao
        r = self.session.get(url)
        time.sleep(1)
        soup = bs(r.content, 'html5lib')

        url = soup.find('div', attrs={'id':'showall'})
        url = url.find('a')['href']
        # print(url)


        r = self.session.get(url)
        time.sleep(1)
        soup = bs(r.content, 'html5lib')
        linhas = soup.find_all('tr',attrs={'class':''})
        for linha in linhas:
            if not 'id' in linha.attrs:
                continue
            cells = linha.find_all('td')
            id = ''
            nome = ''
            email = ''
            papel = ''
            for cell in cells:
                if 'c3' in cell['class']:
                    try:
                        papel = cell.find('a').contents[0]
                    except:
                        papel = cell.contents[0]

                if 'c1' in cell['class']:
                    a = cell.find('a')
                    id = a['href']
                    id = id[id.find('?id=')+4:id.find('&course=')]
                    nome = a.contents[1]
                if 'c2' in cell['class']:
                    email = cell.contents[0]
            if('Estudante' in papel):
                alunos.append(Aluno(nome, True, 0, id, email))

        return alunos

    def obter_atividades_disciplina_moolde(self, disciplina):
        self.log('OBTENDO ATIVIDADES DO CURSO NO MOODLE')
        r = self.session.get(disciplina.get_url_livro_notas())
        soup = bs(r.content, 'html5lib')

        tr = soup.find('tr', attrs={'class': 'heading'})
        ths = tr.find_all('th', attrs={'class': 'item'})

        atividades = []
        for th in ths:
            dataitem = th['data-itemid']
            a = th.find_all('a')[0]
            link = a['href']
            id = link[link.find('?id=')+4:]
            tipo = a.find('img')['alt']
            nome = a.contents[1]
            # print(dataitem, tipo, id, link, nome)

            atividades.append(Atividade(nome, id, tipo, link, dataitem))

        for atividade in atividades:
            if 'Tarefa' in atividade.tipo:
                r = self.session.get(atividade.link)
                time.sleep(1)
                soup = bs(r.content, 'html5lib')
                trs = soup.find('table', attrs={'class': 'generaltable'}).find_all('tr')

                for tr in trs:
                    th = tr.find_all('th')[0].contents[0]
                    td = tr.find_all('td')[0].contents[0]
                    if 'Precisa de avaliação' in th:
                        atividade.precisa_avaliacao = td
                    if 'Data de entrega' in th:
                        atividade.data_entrega = td

        disciplina.ATIVIDADES = atividades
        disciplina.num_atividades = len(disciplina.ATIVIDADES)
        return atividades

    def aluno_que_n_fez_atividade(self, atividade):
        # identificar qual aluno não fez a atividade
        try:
            r = self.session.get(atividade.link + '&action=grading')
        except:
            self.session = self.login(USERNAME, PASSWORD)
            r = self.session.get(atividade.link + '&action=grading')

        time.sleep(1)
        soup = bs(r.content, 'html5lib')
        trs = soup.find_all('tr')
        alunos = []

        for tr in trs:
            if 'class' in tr.attrs and 'user' in tr['class'][0]:
                id_aluno = tr['class'][0]
                td = tr.find_all('td', attrs={'class': 'c4'})[0]
                divs = td.find_all('div')

                status = ''
                for div in divs:
                    if 'submissiongraded' in div['class'][0]:
                        status = 'corrigido'
                        continue
                    if 'submissionstatussubmitted' in div['class'][0]:
                        status = 'enviado'
                        continue
                    if 'submissionstatus' in div['class'][0]:
                        status = 'nao_enviado'

                if status == 'enviado':
                    alunos.append(self.obter_aluno_por_id(id_aluno[4:]))
        return alunos

    def obter_aluno_por_id(self, id):
        for aluno in self.ALUNOS_MOODLE:
            if aluno.id == id:
                return aluno
        return Aluno()

    def atividades_para_corrigir(self, disciplina, planilha):
        for atividade in disciplina.ATIVIDADES:
            if not atividade.precisa_avaliacao == '0' and 'Tarefa' in atividade.tipo:
                alunos = self.aluno_que_n_fez_atividade(atividade)
                for aluno in alunos:
                    try:
                        # cell1 = planilha.find(aluno.nome_planilha)
                        cell2 = planilha.find(atividade.get_nome_curto())
                        link_atividade = baseurl +'mod/assign/view.php?id='+atividade.id+'&rownum=0&action=grader&userid='+aluno.id
                        planilha.update_cell(int(aluno.linha) + 4, cell2.col, '=HYPERLINK("'+link_atividade+'";"'+aluno.get_tutor()+'")')
                        time.sleep(1)
                    except:
                        self.log(f'ALUNO {aluno.nome} FORA DA PLANILHA')
                        print(aluno.linha)

    def obter_notas_alunos(self, disciplina):
        self.log("OBTENDO NOTA DOS ALUNOS NO MOODLE")
        r = self.session.get(disciplina.get_url_livro_notas())
        time.sleep(1)
        soup = bs(r.content, 'html5lib')
        notas_alunos = {}
        for aluno in self.ALUNOS_MOODLE:
            notas = []
            for atividade in disciplina.ATIVIDADES:
                id = 'u'+aluno.id+'i'+atividade.dataitem
                td = soup.find('td', attrs={'id':id})
                try:
                    sp = td.find('span')
                    nota = sp.contents[0]
                except:
                    nota = '-'
                # print(nota, aluno.nome, atividade.nome)
                notas.append(nota)
            notas_alunos[aluno.nome] = notas
        return notas_alunos


script = Script()
script.iniciar()