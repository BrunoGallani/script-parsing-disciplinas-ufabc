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

        # 4. Leitura dos PDFs de matrículas e das planilhas de planejamento da pasta de arquivos a analisar
        input_dir = 'arquivos_analise'

        if not os.path.exists(input_dir):
            os.makedirs(input_dir)

        pdfs = []
        planilhas = []
        quad_pattern = re.compile(r'\d{4}\.[1-3]')

        for arquivo in sorted(os.listdir(input_dir)):
            path_arquivo = os.path.join(os.getcwd(), input_dir, arquivo)
            extension = os.path.splitext(arquivo)[1].lower()

            if arquivo.startswith('~$'):
                continue

            if extension == '.pdf':
                pdfs.append(path_arquivo)

            elif extension == '.xlsx':
                quad_match = quad_pattern.search(arquivo)

                if not quad_match:
                    raise ValueError(f'Quadrimestre não identificado no nome do arquivo "{arquivo}" (exemplo: "Turmas 2027.3.xlsx")')

                planilhas.append((path_arquivo, quad_match.group()))

        if not pdfs and not planilhas:
            raise FileNotFoundError(f'Nenhum PDF de matrículas ou planilha de planejamento encontrado na pasta "{input_dir}"')

        # 5. Parsing do PDF e/ou da planilha para extrair as disciplinas disponíveis

        outputs_dir = 'analise_disciplinas'

        if not os.path.exists(outputs_dir):
            os.makedirs(outputs_dir)

        def log(text):
            print(str(text))
            with open(LOG_FILE, 'a') as f:
                f.write(str(text) + '\n')

        analises = []

        for path_pdf in pdfs:
            pdf_matriculas_disciplinas = PdfReader(path_pdf)

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

            analises.append(('pdf de matrículas', 'pdf_matriculas', quad_number, disciplinas_disponiveis))

        for path_planilha, quad_planilha in planilhas:
            df_turmas = pd.read_excel(path_planilha, usecols=['Sigla disciplina'], engine='openpyxl')

            disciplinas_planilha = list(dict.fromkeys(
                df_turmas['Sigla disciplina'].dropna().astype(str).str.strip()
            ))

            analises.append(('planilha de planejamento', 'planilha_planejamento', quad_planilha, disciplinas_planilha))

        for origem, sufixo_arquivo, quad_number, disciplinas_disponiveis in analises:
            filename = f'{sufixo_arquivo}_{quad_number.replace(".", "-")}.txt'
            LOG_FILE = os.path.join(outputs_dir, filename)

            if os.path.exists(LOG_FILE):
                os.remove(LOG_FILE)

            log(f"Disciplinas disponíveis - quadrimestre {quad_number} ({origem}):\n")
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

        pag.alert('''Verifique os resultados no terminal e na pasta \"analise_disciplinas\".''',
                title = 'Fim da automação')
    
    except Exception as e:
        pag.alert(f'Ocorreu um erro durante a execução do script:\n{e}',
                title = 'Erro na automação')

if "__main__" == __name__:
    main()