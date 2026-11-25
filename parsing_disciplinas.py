import pandas as pd
import os
import re
from PyPDF2 import PdfReader
import pyautogui as pag

def main():

    try:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        # 1. Análise de disciplinas - BCC

        # 1.1 Obrigatórias
        df_obrigatorias_bcc = pd.read_excel("disciplinas_bacharelado_bcc.xlsx",
                                        sheet_name="Disciplinas obrigatórias")

        obrigatorias_bcc = {
            sigla.strip(): disciplina.strip() 
            for sigla, disciplina in df_obrigatorias_bcc[['Sigla', 'Disciplina']].values
        }
        
        # 1.2 Limitadas
        df_limitadas_bcc = pd.read_excel("./disciplinas_bacharelado_bcc.xlsx", sheet_name="Disciplinas limitadas", usecols="A:E", engine='openpyxl')

        limitadas_bcc = {
            sigla.strip(): disciplina.strip() 
            for sigla, disciplina in df_limitadas_bcc[['Sigla', 'Disciplina']].values
        }

        # 2. Análise de disciplinas - BCD

        # 2.1 Obrigatórias
        df_obrigatorias_bcd = pd.read_excel("disciplinas_bacharelado_bcd.xlsx",
                                        sheet_name="Disciplinas obrigatórias")

        obrigatorias_bcd = {
            sigla.strip(): disciplina.strip()
            for sigla, disciplina in df_obrigatorias_bcd[['Sigla', 'Disciplina']].values
        }

        # 2.2 Limitadas
        df_limitadas_bcd = pd.read_excel("./disciplinas_bacharelado_bcd.xlsx", sheet_name="Disciplinas limitadas", usecols="A:E", engine='openpyxl')

        limitadas_bcd = {
            sigla.strip(): disciplina.strip()
            for sigla, disciplina in df_limitadas_bcd[['Sigla', 'Disciplina']].values
        }

        # 3. Análise de disciplinas limitadas compartilhadas entre BCC e BCD
        df_limitadas_compartilhadas_bcc_bcd = pd.read_excel("./disciplinas_bacharelado_bcc.xlsx",
                                                            sheet_name="Disciplinas limitadas BCD e BCC", engine='openpyxl')

        limitadas_bcc_bcd = {
            sigla.strip(): disciplina.strip()
            for sigla, disciplina in df_limitadas_compartilhadas_bcc_bcd[['Sigla', 'Disciplina']].values
        }

        # 4. Leitura do PDF com as disciplinas disponíveis no quad vigente
        for arquivo in os.listdir():
            extension_pattern = re.compile(r'.+(\.[a-z]+$)')
            extension_match = extension_pattern.search(arquivo)
            extension = extension_match.group(1) if extension_match else ''

            if extension == '.pdf':
                path_arquivo = os.path.join(os.getcwd(), arquivo)
                break

        pdf_matriculas_disciplinas = PdfReader(path_arquivo)

        # 5. Parsing do PDF para extrair as disciplinas limitadas disponíveis

        outputs_dir = 'outputs'
        filename = 'analise_disciplinas.txt'
        LOG_FILE = os.path.join(outputs_dir, filename)

        if not os.path.exists(outputs_dir):
            os.makedirs(outputs_dir)

        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)

        def log(text):
            print(str(text))
            with open(LOG_FILE, 'a') as f:
                f.write(str(text) + '\n')

        # Captura das disciplinas
        disciplinas = []
        for page_number in range(len(pdf_matriculas_disciplinas.pages)):
            pagina_atual = pdf_matriculas_disciplinas.pages[page_number]
            texto = pagina_atual.extract_text()

            disciplina_pattern = re.compile(r'[DN][ABCDEI][0-9]*([A-Z]{3,4}[0-9]{3,4}\-\d{2,})(SA|SB)',
                                            re.MULTILINE | re.IGNORECASE)

            for qtd_disciplinas in range(len(disciplina_pattern.findall(texto))):
                disciplinas.append(disciplina_pattern.findall(texto)[qtd_disciplinas][0].replace('\n', '').strip())

        disciplinas_disponiveis = []
        for disciplina in disciplinas:
            if disciplina not in disciplinas_disponiveis:
                disciplinas_disponiveis.append(disciplina)
        
        # Captura do quadrimestre atual
        for page_number in range(len(pdf_matriculas_disciplinas.pages)):
            pagina_atual = pdf_matriculas_disciplinas.pages[page_number]
            texto = pagina_atual.extract_text()

            quad_number_pattern = re.compile(r'\d{4}\.[1-3]')
            quad_number = quad_number_pattern.findall(texto)[0]
            break
        
        log(f"Disciplinas disponíveis - quadrimestre {quad_number}:\n")
        log("-------------------------------- CIÊNCIA DA COMPUTAÇÃO --------------------------------")
        log(f'\n{"-"*10} OBRIGATÓRIAS {"-"*10}\n')
        for disciplina in disciplinas_disponiveis:
            if obrigatorias_bcc.get(disciplina):
                log(obrigatorias_bcc.get(disciplina))

        log(f'\n{"-"*10} LIMITADAS {"-"*10}\n')
        for disciplina in disciplinas_disponiveis:
            if limitadas_bcc.get(disciplina):
                log(limitadas_bcc.get(disciplina))

        log("\n-------------------------------- CIÊNCIA DE DADOS --------------------------------")
        log(f'\n{"-"*10} OBRIGATÓRIAS {"-"*10}\n')
        for disciplina in disciplinas_disponiveis:
            if obrigatorias_bcd.get(disciplina):
                log(obrigatorias_bcd.get(disciplina))

        log(f'\n{"-"*10} LIMITADAS {"-"*10}\n')
        for disciplina in disciplinas_disponiveis:
            if limitadas_bcd.get(disciplina):
                log(limitadas_bcd.get(disciplina))

        log("\n-------------------------------- LIMITADAS DE AMBOS OS CURSOS --------------------------------\n")
        for disciplina in disciplinas_disponiveis:
            if limitadas_bcc_bcd.get(disciplina):
                log(limitadas_bcc_bcd.get(disciplina))

        print("-" * 10)
        print(f'Arquivo salvo em: {os.path.abspath(LOG_FILE)}')

        pag.alert('''Verifique os resultados, e quando terminar, clique em "ok".''',
                title = 'Fim da automação')
    
    except Exception as e:
        pag.alert(f'Ocorreu um erro durante a execução do script:\n{e}',
                title = 'Erro na automação')

if "__main__" == __name__:
    main()