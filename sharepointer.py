import time
import getpass
from pathlib import Path
from pathlib import WindowsPath
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

class SharepointerSelenium:
    
    def __init__(self, sharepoint_url: str, username: str = None, password: str = None, download_dir: str = None):
        '''
        # Descrição
        Objeto de classe para acesso e manipulação de documentos no Sharepoint.

        Parameters
        ----------
        sharepoint_url : string
            - Endereço base do sharepoint
            - **exemplo de parametro:**
            https://nomedosharepoint.sharepoint.com

        username : string (opcional)
            - Seu e-mail de usuário
            - Na não definição, o código perguntará no momento do login

        password : string (opcional)
            - Sua senha 
            - Recomenda-se a utilização de variaveis de ambiente
            para preservar a integridade de sua credencial
            - Na não definição, o código perguntará no momento do login

        download_dir : string (opcional)
            - Caminho onde serão salvados os arquivos baixados
            - Na não definição, os arquivos serão baixados na pasta 'downloads'
            interna ao diretório deste código
        '''        
        self._username = username
        self._password = password
        self._sharepoint_url = sharepoint_url.rstrip('/')
        self._logged = False
        self._download_dir = str(Path(__file__).parent / 'downloads') if download_dir is None else download_dir

        # Coletando e validando webdriver edge
        webdriver_dir = Path(__file__).parent / 'webdriver'
        self._edge_webdriver_path = str(next((f for f in webdriver_dir.glob('*.exe') if 'edge' in f.name.lower()), None))

        if self._edge_webdriver_path is None:
            raise FileNotFoundError('Nenhum webdriver "edge" foi encontrado no diretório "webdriver"')

        # Gerando driver
        self._driver = self._initizalize_edge_drive()

        # Configurando botões
        self._initialize_buttons_and_elements()

    @property
    def username(self):
        if self._username == None:
            self._username = input('Insira seu e-mail: ')
        return self._username
    
    @username.setter
    def username(self, username: str):
        self._username = username
    
    @property
    def password(self):
        if self._password == None:
            self._password = getpass.getpass('Insira sua senha: ')
        return self._password
    
    @password.setter
    def password(self, password: str):
        self._password = password

    # Configuração do objeto
    # ----------------------------------------------------------------------
    # Métodos do objeto

    def download_file(self, sharepoint_folder_url: str, file_name_with_ext: str, wait_download_time: int = 60) -> None:
        '''
        # Descrição
        Realiza o download de um arquivo desejado dentro de um Sharepoint
        específico

        Parameters
        ----------
        sharepoint_folder_url : string
            - Endereço da pasta Sharepoint onde o arquivo está localizado
            - Para visualizar o endereço, clique com o botão direito na
            pasta do Sharepoint o qual deseja acessar o documento

        file_name_with_ext : string
            - Nome do arquivo com sua extensão
            - Exemplo: 'exemplo.xlsx'
        
        wait_download_time : int (opcional)
            - Tempo máximo de espera para realização do download
            em segundos
            - Por padrão, este valor estará como 60 segundos
            - Caso passe deste tempo, o código entenderá que houve
            algum erro
            - Ajuste este valor em caso de arquivos muito pesados
            que possam levar mais de 60 segundos, caso contrário,
            a aplicação interromperá o donwload
        '''
        if self._logged == False:
            self._login()

        self._driver.get(sharepoint_folder_url)

        checkbox_button = _ElementClickXpath(f"//div[@role='checkbox' and contains(@aria-label, '{file_name_with_ext}')]//div", self._driver)
        checkbox_button.click()
        
        file_name_with_ext = self._increment_file_name(file_name_with_ext)

        try:
            self._MORE_BUTTON.click()
        except TimeoutException:
            print('Botão more não encontrado')
            pass

        self._DOWNLOAD_BUTTON.click()

        if self._wait_for_file_download(file_name_with_ext, wait_download_time):
            print(f'Download do arquivo {file_name_with_ext} concluído.')
        else:
            print(f'Tempo esgotado para o download do arquivo {file_name_with_ext}')

    def upload_file(self, local_files_path: str | WindowsPath | list, sharepoint_folder_url: str, replace: bool = None, max_attempts: int = 60) -> bool:
        '''
        # Descrição
        Realiza o upload de um ou mais arquivos a um Sharepoint específico

        Parameters
        ----------
        local_files_path : string | list
            - Endereço da pasta local onde o(s) arquivo(s) está(ão)
            localizado(s)
            - Em caso de multiplo arquivo, dever-se-á passar os endereços
            como uma lista de string

        sharepoint_folder_url : string
            - Endereço da pasta Sharepoint onde será realizado o upload
            - Para visualizar o endereço, clique com o botão direito na
            pasta do Sharepoint o qual deseja acessar o documento
        
        replace : bool (opcional)
            - Pré estabelece se, ao encontrar um arquivo com
            mesmo nome, sobrescreverá ou manterá ambos os arquivos
            - O programa perguntará se deseja ou não sobrescrever
            caso não definido este parâmetro
            - **True**: Sobrescreva
            - **False**: Mantenha ambos

        max_attempts : int (opcional)
            - Quantidade de tentativas para encontrar o alerta de
            sucesso no upload ou no aviso de arquivo com mesmo nome
            - Por padrão, este valor estará como 60 tentativas
            - Caso realizar todas as tentativas sem sucesso em encontrar
            os alertas, o código entenderá que houve algum erro
            - Ajuste este valor em caso de arquivos muito pesados
            que possam levar mais de 60 ciclos, mas, este número deve
            atender a grande maioria das situações

        Returns
        ----------
        True
            - Operação realizada com sucesso

        False
            - Operação não realizada
        '''
        if self._logged == False:
            self._login()
        
        self._driver.get(sharepoint_folder_url)
        self._UPLOAD_BUTTON.click()
        self._UPLOAD_FILE_BUTTON.click()
        # NOTE O campo input só aparece após clicar no upload files
        # A janela abrirá, mas o upload será feito diretamente pelo html
        upload_element = self._wait_by_xpath(f"//input[@type='file']")

        def _get_absolute_path(path):
            return str(Path(path).resolve())

        def _configure_upload_settings(local_files_path: str | WindowsPath | list) -> tuple:
            if isinstance(local_files_path, (str, WindowsPath)) or (isinstance(local_files_path, list) and len(local_files_path) == 1):
                local_files_path = _get_absolute_path(local_files_path[0] if isinstance(local_files_path, list) else local_files_path)
                success_xpath = self._SUCCESS_SINGLE_ALERT_XPATH
                replace_button = self._REPLACE_FILE_BUTTON
                keep_button = self._KEEP_BOTH_FILE_BUTTON
            else:
                local_files_path = [_get_absolute_path(path) for path in local_files_path]
                success_xpath = self._SUCCESS_MULTIPLE_ALERT_XPATH
                replace_button = self._REPLACE_ALL_BUTTON
                keep_button = self._KEEP_ALL_FILE_BUTTON
            
            return local_files_path, success_xpath, replace_button, keep_button

        local_files_path, success_xpath, replace_button, keep_button = _configure_upload_settings(local_files_path)
        upload_element.send_keys(local_files_path if isinstance(local_files_path, str) else '\n'.join(local_files_path))

        for attempt in range(max_attempts):
            try:
                # Verifica se o alerta de arquivo existente aparece
                self._wait_by_xpath(self._FILE_EXISTS_ALERT_XPATH, time_out=3)

                if replace is None:
                    replace = self._input_replace()

                if replace:
                    print("Substituindo o arquivo...")
                    replace_button.click()
                else:
                    print("Mantendo o arquivo existente...")
                    keep_button.click()

                # Após lidar com o alerta, aguarda o upload do arquivo
                if self._wait_for_file_upload(success_xpath, 1000):
                    print("Upload do(s) arquivo(s) concluído(s)")
                    return True
                
            except Exception:
                # Se o alerta não aparecer, verifica se o upload foi concluído diretamente
                if self._wait_for_file_upload(success_xpath, 3):
                    print("Upload do(s) arquivo(s) concluído(s)")
                    return True      
                else:
                    pass

            # Se nenhum elemento foi encontrado e o loop ainda está em andamento, continua tentando
            print(f"#{attempt + 1} - Buscando: Aguardando alerta de sucesso ou arquivo de mesmo nome...")

        print(f"Máximo de tentativas ({max_attempts}) alcançado. Upload falhou.")
        
        return False

    # Métodos do objeto
    # ----------------------------------------------------------------------
    # Métodos internos do objeto

    def _initizalize_edge_drive(self):
        edge_options = Options()
        edge_options.use_chromium = True  # Necessário para Edge Chromium
        edge_options.add_argument("--inprivate")
        # Suprime log, exibe apenas erros fatais
        edge_options.add_argument('--log-level=3')

        # Configuração de download
        prefs = {
            "download.default_directory": self._download_dir,
            "profile.default_content_settings.popups": 0,
            "directory_upgrade": True
        }
        edge_options.add_experimental_option("prefs", prefs)
        
        # Gerando driver
        return webdriver.Edge(service=Service(self._edge_webdriver_path), options=edge_options)
    
    def _initialize_buttons_and_elements(self):
        self._DOWNLOAD_BUTTON = _ElementClickXpath(f'//button[@data-automationid="downloadCommand"]', self._driver)
        self._UPLOAD_BUTTON = _ElementClickXpath(f'//button[@data-automationid="uploadCommand"]', self._driver)
        self._UPLOAD_FILE_BUTTON = _ElementClickXpath(f'//span[contains(@class, "ms-ContextualMenu-itemText") and text()="Files"]', self._driver)
        self._MORE_BUTTON = _ElementClickXpath(f'//i[@data-icon-name="More"]', self._driver)
        self._KEEP_BOTH_FILE_BUTTON = _ElementClickXpath(f'//span[contains(@class, "ms-Button-flexContainer") and text()="Keep both"]', self._driver)
        self._KEEP_ALL_FILE_BUTTON = _ElementClickXpath(f'//span[contains(@class, "ms-Button-flexContainer") and text()="Keep all"]', self._driver)
        self._REPLACE_FILE_BUTTON = _ElementClickXpath(f'//span[contains(@class, "ms-Button-flexContainer") and text()="Replace"]', self._driver)
        self._REPLACE_ALL_BUTTON = _ElementClickXpath(f'//span[contains(@class, "ms-Button-flexContainer") and text()="Replace all"]', self._driver)
        self._SUCCESS_SINGLE_ALERT_XPATH = '//label[contains(@class, "od-Notify-message")]'
        self._SUCCESS_MULTIPLE_ALERT_XPATH = '//div[contains(@class, "title_7bf3db39")]'
        self._FILE_EXISTS_ALERT_XPATH = '//span[contains(@class, "ms-Button-flexContainer") and contains(text(), "Keep")]'

    def _login(self):
        self._driver.get(self._sharepoint_url)
        
        # Loop até e-mail valido
        while not self._insert_email():
            print(f'E-mail {self.username} não existe!')
            self.username = None

        print(f'E-mail {self.username} inserido com sucesso!')
        
        # Loop até senha valida
        while not self._insert_password():
            print(f'Senha incorreta!')
            self.password = None

        # Tela Stay signed in? | Continuar conectado?
        try:
            no_stay_signed_btn = self._wait_by_id('idBtn_Back', 2)
            no_stay_signed_btn.click()
        except:
            pass
        
        try:
            self._wait_by_id('ms-error', 1)
            print(f'\nE-mail {self.username} não possui autorização para acessar este sharepoint')
            print('Gostaria de inserir outro e-mail e senha?')
            continuar = input('Digite "s" para continuar, ou qualquer outra tecla para encerrar: ')
            
            if continuar.lower() == 's':
                self._reset_session()
                return
            else:
                raise KeyboardInterrupt("Sharepointer encerrado pelo usuário")
        except TimeoutException:
            pass

        self._logged = True
        print('Login realizado com sucesso')

    def _insert_email(self) -> bool:
        email_input = self._wait_by_element_name('loginfmt', 20)
        email_input.clear() # Limpa cache do email
        email_input.send_keys(self.username)
        email_input.send_keys(Keys.ENTER)
        time.sleep(.5) # Tempo para não capturar mensagem existente
        try:
            self._wait_by_id('usernameError', .5)
            return False
        except:
            return True
        
    def _insert_password(self) -> bool:
        password_input = self._wait_by_element_name('passwd', 20)
        password_input.clear() # Limpa cache da senha
        password_input.send_keys(self.password)
        time.sleep(.5) # Garantir inserção da senha
        password_input.send_keys(Keys.ENTER)
        time.sleep(.5) # Tempo para não capturar mensagem existente
        try:
            self._wait_by_id('passwordError', .5)
            return False
        except:
            return True
        
    def _reset_session(self):
        print('Reinicializando sessão...')
        self.username = None
        self.password = None
        self._logged = False

        try:
            self._driver.quit()
        except Exception as e:
            print(f"Erro ao fechar o driver: {e}")

        self._driver = self._initizalize_edge_drive()
        self._initialize_buttons_and_elements()
        self._login() 

    def _wait_by_id(self, id_name: str, time_out: int = 10):
        return WebDriverWait(self._driver, time_out).until(
                EC.presence_of_element_located((By.ID, id_name))
            )

    def _wait_by_element_name(self, el_name: str, time_out: int = 10):
        return WebDriverWait(self._driver, time_out).until(
                EC.presence_of_element_located((By.NAME, el_name))
            )
    
    def _wait_by_xpath(self, xpath: str, time_out: int = 10):
        return WebDriverWait(self._driver, time_out).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

    def _wait_for_file_download(self, file_name: str, time_out: int):
        local_file_path = Path(f'{self._download_dir}/{file_name}')
        # time.time() momento atual em segundos
        end_time = time.time() + time_out

        while time.time() < end_time:
            if local_file_path.exists():
                return True
            time.sleep(1)

        return False
    
    def _wait_for_file_upload(self, success_xpath: str, time_out: int):
        # time.time() momento atual em segundos
        end_time = time.time() + time_out

        while time.time() < end_time:
            try:
                self._wait_by_xpath(success_xpath)
                return True
            except:
                pass
            time.sleep(1)

        return False
    
    def _increment_file_name(self, file_name: str):
        name, ext = file_name.rsplit('.', 1)
        files_with_same_name = [f for f in Path(self._download_dir).glob(f'{name}*.{ext}')]
        
        if files_with_same_name:
            file_name = f'{name} ({len(files_with_same_name)}).{ext}'

        return file_name
    
    def _input_replace(self) -> str:
        while True:
            user_input = input(f'O(s) arquivo(s) já existe(m). Deseja sobrescrever? (s/n): ')
            if user_input.lower() == 's':
                replace = True
                break
            elif user_input.lower() == 'n':
                replace = False
                break
            else:
                print('Resposta inválida!!!')

        return replace

# Métodos internos do objeto
# ----------------------------------------------------------------------
# Definição da classe interna _ElementClickXpath

class _ElementClickXpath:
    def __init__(self, xpath: str, driver: webdriver):
        self._xpath = xpath
        self._driver = driver

    def click(self, time_out: int = 10) -> bool:
        try:
            element = self._find_element(time_out)
            element.click()
            return True
        except TimeoutException:
            print(f"Elemento com XPath '{self._xpath}' não foi encontrado no tempo definido.")
            return False
        except StaleElementReferenceException:
            try:
                print("Elemento não encontrado, tentando encontrar o elemento novamente...")
                time.sleep(1)
                element = self._find_element(time_out)
                element.click()
                return True
            except Exception as e:
                print(f"Erro ao tentar clicar no elemento novamente: {str(e)}")
                return False
            
    def _find_element(self, time_out: int = 10):
        return WebDriverWait(self._driver, time_out).until(
            EC.presence_of_element_located((By.XPATH, self._xpath))
        )