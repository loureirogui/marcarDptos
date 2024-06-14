import traceback
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from docx import Document
import openpyxl

# Solicita as credenciais do usuário
emailLogin = input("Qual seu email de login no acessórias?\n")
senhaLogin = input("Qual sua senha de login no acessórias?\n")
print("Marcando os departamentos nas empresas. Aguarde por gentileza...")

# Caminho para o driver do Edge
driver_path = 'msedgedriver.exe'

# Configura as opções do Edge
edge_options = Options()
edge_options.headless = False  # Executa o Edge em modo não headless

# Inicializa o navegador Edge
service = Service(driver_path)
edge_driver = webdriver.Edge(service=service, options=edge_options)

# Abre o link desejado
url = f"https://app.acessorias.com/sysmain.php?m=105&act=e&i=365&uP=14&o=EmpNome,EmpID|Asc"
edge_driver.get(url)

# Carrega o arquivo .xlsx
workbook = openpyxl.load_workbook('empresas.xlsx')

# Seleciona a planilha ativa (a primeira planilha aberta por padrão)
sheet = workbook.active

# Seleciona a coluna específica (por exemplo, coluna B)
coluna = 'B'

# Lógica de login
try:
    # Espera o campo de e-mail aparecer
    email_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'mailAC'))
    )
    # Insere o e-mail no campo
    email_input.send_keys(emailLogin)
except Exception as e:
    print("Erro ao inserir o e-mail no campo de login:", str(e))
    traceback.print_exc()

try:
    # Espera o campo de senha aparecer
    senha_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'passAC'))
    )
    # Insere a senha no campo
    senha_input.send_keys(senhaLogin)
except Exception as e:
    print("Erro ao inserir a senha no campo de senha:", str(e))
    traceback.print_exc()

# Espera o botão de login aparecer
try:
    login_button = WebDriverWait(edge_driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.rounded.large.expanded.primary-degrade.btn-enviar'))
    )
    # Clique no botão de login
    login_button.click()
except Exception as e:
    print("Erro ao clicar no botão de login:", str(e))
    traceback.print_exc()

# Espera 2 segundos após clicar no botão de login
time.sleep(2)

idEmpresa = ''
for cell in sheet[coluna][3:]:
    idEmpresa = cell.value
    if idEmpresa:
        try:
            url = f"https://app.acessorias.com/sysmain.php?m=105&act=e&i={idEmpresa}&uP=14&o=EmpNome,EmpID|Asc"
            edge_driver.get(url)

            time.sleep(2)

            # Espera até o ícone do grupo estar clicável
            contact_icon_element = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'iDivCtt'))
            )

            # Obtém a classe atual do ícone do grupo
            contact_icon_class = contact_icon_element.get_attribute("class")

            # Se a classe contém 'grey', clica no ícone para mudar para 'green'
            if 'grey' in contact_icon_class:
                contact_icon_element.click()

            try:
                # Espera até que todos os elementos do botão roxo estejam presentes
                WebDriverWait(edge_driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button.btn.btn-sm.btn-purple[type="button"][title="Selecionar departamentos"]'))
                )

                # Busca todos os elementos correspondentes
                deptos_icon_elements = edge_driver.find_elements(By.CSS_SELECTOR, 'button.btn.btn-sm.btn-purple[type="button"][title="Selecionar departamentos"]')

                # Itera a partir do segundo elemento (índice 1)
                for i in range(1, len(deptos_icon_elements)):
                    try:
                        # Rebusca todos os elementos a cada iteração para garantir que a lista esteja atualizada
                        deptos_icon_elements = edge_driver.find_elements(By.CSS_SELECTOR, 'button.btn.btn-sm.btn-purple[type="button"][title="Selecionar departamentos"]')
                        deptos_icon_element = deptos_icon_elements[i]
                        deptos_icon_element.click()

                        try:
                            # Rebusca todos os elementos de editar contato a cada iteração para garantir que a lista esteja atualizada
                            edit_buttons = edge_driver.find_elements(By.CSS_SELECTOR, 'button.btn.btn-sm.btn-yellow.col-xs-4.col-sm-4[title="Editar dados do contato"]')
                            edit_button = edit_buttons[i - 1]  # Clique no botão amarelo de índice i-1
                            edit_button.click()
                        except Exception:
                            print("Erro ao clicar no botão de editar contato:")
                            traceback.print_exc()

                        try:
                            # Incrementa o índice do checkbox dinamicamente
                            checkbox_id = f'0_{i}'
                            check_button = WebDriverWait(edge_driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, f'a[onclick="selectAllDptos(\'{checkbox_id}\', true);"]'))
                            )
                            check_button.click()
                        except Exception:
                            print("Erro ao clicar no botão de check")
                            traceback.print_exc()

                        try:
                            # O índice do botão de salvar permanece constante
                            save_button = WebDriverWait(edge_driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[onclick="saveCtt(\'{idEmpresa}\', \'{checkbox_id}\', \'0\');"]'))
                            )
                            save_button.click()
                            time.sleep(2)
                            deptos_icon_element.click()
                        except Exception:
                            print("Erro ao clicar no botão de salvar:")
                            traceback.print_exc()

                    except Exception:
                        print('Erro ao clicar no botão roxo no índice', i)

            except Exception:
                print("Erro ao buscar os botões roxos:")
                traceback.print_exc()

        except Exception as e:
            print('Não foi possível marcar os departamentos na empresa: ' + str(idEmpresa))
            traceback.print_exc()
