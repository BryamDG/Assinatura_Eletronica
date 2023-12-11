from tkinter import Tk, Label, Button, Entry, OptionMenu, StringVar
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from time import sleep

class InterfaceGrafica:
    def __init__(self, master, outlook_automation):
        self.master = master
        master.title("Automação de Assinatura Eletrônica")
        self.label = Label(master, text="Digite sua senha e preencha os dados da assinatura")
        self.label.pack()

        self.senha_label = Label(master, text="Senha:")
        self.senha_label.pack()
        self.senha = Entry(master, show="*")
        self.senha.pack()

        self.nome_label = Label(master, text="Nome:")
        self.nome_label.pack()
        self.nome = Entry(master)
        self.nome.pack()

        self.cargo_label = Label(master, text="Cargo:")
        self.cargo_label.pack()
        self.cargo = Entry(master)
        self.cargo.pack()

        self.email_label = Label(master, text="E-mail:")
        self.email_label.pack()
        self.email = Entry(master)
        self.email.pack()

        self.ramal_label = Label(master, text="Ramal:")
        self.ramal_label.pack()
        self.ramal = Entry(master)
        self.ramal.pack()      

        self.departamento_label = Label(master, text="Departamento:")
        self.departamento_opcoes = ["CTIC", "Cerimonial", "D.A", "DFO", "GPAO", "Ouvidoria", "Protocolo", "UFC", "UDBL", "UPPH", "DRH", "CJ", "ATGS", "LGPD", "ASSPAR", "Gabinete", "UPPM", "PROAC", "UM"]
        self.departamento_var = StringVar(master)
        self.departamento_var.set(self.departamento_opcoes[0])
        self.departamento_menu = OptionMenu(master, self.departamento_var, *self.departamento_opcoes)
        self.departamento_menu.pack()

        self.iniciar_button = Button(master, text="Iniciar Automação de Assinatura e E-mail", command=self.iniciar_automacao)
        self.iniciar_button.pack()

        self.outlook_automation = outlook_automation

    def iniciar_automacao(self):
        nome = self.nome.get()
        senha = self.senha.get()
        cargo = self.cargo.get()
        email = self.email.get()
        ramal = self.ramal.get()
        departamento = self.departamento_var.get()
        self.outlook_automation.iniciar(nome, senha, cargo, email, ramal, departamento)
        self.label.config(text="Automação concluída com sucesso")


class OutlookAutomation:
    def iniciar(self, nome, senha, cargo, email, ramal, departamento):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument('--lang=pt-BR')
        chrome_options.add_argument('--disable-notifications')

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        navegador.implicitly_wait(10)

        if departamento == "CTIC":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/doc2.aspx?sourcedoc=%7BF7F1FDF9-8274-4469-8BBC-BF004A8B48D4%7D&file=Assinatura%20CTIC.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&ct=1701881420389&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=4cb3d518-d9e1-488b-9eba-d93aacbe3896&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d")

        elif departamento == "D.A":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B58940844-D92C-4D1F-A40A-389B3F08D7B2%7D&file=sas.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1699467309919&wdOrigin=OFFICECOM-WEB.START.REC&cid=8b8741c2-6fa4-4316-aec6-052c2cbe9a37&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=fd65b62a-ae69-4518-8b39-69373256cd02")

        elif departamento == "DFO":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B91B7B134-6237-4F97-937B-2C9C2F4D615B%7D&file=Cerimonial.odp&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881464285&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=c5bdac44-c321-46c2-9fa5-308fa860ff46&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "UPPH":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B17501114-7BC2-40EA-BB03-E14474157525%7D&file=Assinatura%20UPPH.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881679469&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=92efe5b3-fad6-461c-8f65-64bb9f80aa25&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")      

        elif departamento == "UPPM":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B52B552AD-5F21-48EC-9D31-13C84FDA8C41%7D&file=Assinatura%20UPPM.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881763438&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=5600d666-10be-4c32-a883-6586b884938a&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "UFC":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B53D25B85-A408-4175-9005-74D4EE349CDB%7D&file=Assinatura%20UFC.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881748868&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=1d0a3264-98b7-470b-8462-90260a5844d4&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "Ouvidoria":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B17699334-BD37-4486-9E34-85F54B3ED13F%7D&file=Assinatura%20Ouvidoria.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881728619&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=34de6e95-d686-4a49-ab94-7fd1b619bea0&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "DRH":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B9C4BF420-15FD-4D04-9CA7-2B439FAA47E2%7D&file=Assinatura%20DRH.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881697694&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=42c7e967-4399-4c7c-956c-8b214d9f813a&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")      

        elif departamento == "ATGS":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7BFD984475-DD5C-496E-9B87-9650BFDD6FAD%7D&file=Pontos%20de%20Cultura.odp&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881445419&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=4333b744-7d37-4b14-89cd-80d672b36865&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "CJ":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7BE61DCA31-DBD4-4133-AA7E-B4D51D82003D%7D&file=Assinatura%20CJ.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881644869&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=efdcaed4-97f3-401d-b840-5748bd6d5d91&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "Cerimonial":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B002B0D0D-4C13-4327-B201-CC2072FBF387%7D&file=Assinatura%20Cerimonial.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881596461&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=e6a40e91-fee9-4047-82ce-d3f3af03028b&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "Gabinete":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B87FB5FB6-2FDE-44D7-8275-6BEB02F0E052%7D&file=Assinatura%20Gabinete.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881483773&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=10d0afea-26ff-4804-91c9-22e244a1d3a3&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")      

        elif departamento == "Protocolo":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B65506F6B-A8C4-4E2E-905F-359C1D70B225%7D&file=Assinatura%20PROTOCOLO.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881710723&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=2a725431-7fb1-41cc-8fe4-8110cab9f16f&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "UDBL":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B865D9213-90A8-4799-BA6B-8006C2B729B7%7D&file=Alessandra%20Borsato%20de%20Oliveira%20Paula.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881950086&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=6c5bab23-cafc-42b2-b83b-83703e18563e&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "PROAC":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/suporte_ctic_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B971902C2-F09A-41A9-8DA9-F895C56112C7%7D&file=Assinatura%20PROAC.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881812716&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=e6a4b11d-5b3a-4387-9539-8d936c5a027e&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")

        elif departamento == "UM":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7BD650A7F2-7B24-4B69-8466-13F7602368F8%7D&file=Assinatura%20UM.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881775963&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=b3dda924-ee68-4808-8779-23f5a0d4d755&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d ")      

        elif departamento == "ASSPAR":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7BD0EED5C5-61EC-47DC-81B7-F2FFD970C745%7D&file=Assinatura%20ASSPAR.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701881523533&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=2bedd81e-3a0e-48d1-ad26-ba2863cfd150&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d")
                    
        elif departamento == "GPAO":
            navegador.get("https://governosp-my.sharepoint.com/:p:/r/personal/mgdsantos_sp_gov_br/_layouts/15/Doc.aspx?sourcedoc=%7B5CC37D88-EE15-4391-B19C-73BCDC1B7410%7D&file=Angela%20Harumi%20Uechi.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1&login_hint=suporte_ctic%40sp.gov.br&ct=1701882174478&wdOrigin=OFFICECOM-WEB.START.EDGEWORTH&cid=2c9c07e1-8867-4b4a-ba6a-4600d3c9148e&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=70dc8260-3aad-47ba-a160-7d2c01c7151d")

        sleep(2)
        senha_element = navegador.find_element(By.XPATH, '//*[@id="i0118"]')
        senha_element.send_keys(senha)
        entrar_button = navegador.find_element(By.XPATH, '//*[@id="idSIButton9"]')
        entrar_button.click()
        sleep(4)

        largura, altura=pyautogui.size()
        x_meio=largura/2.5
        y_meio=altura/2
        pyautogui.click(x_meio, y_meio)

        complemento = f'| 11 3339-{ramal}'
        mensagem = f"""Bom Dia!! Conforme o novo decreto N° 67.765 do governo, segue em anexo o manual e a assinatura padrao, siga o manual para inserir a assinatura no Outlook.
        """

        pyautogui.press('tab')
        pyautogui.press('menu')
        pyautogui.press('enter')
        sleep(3)

        pyautogui.hotkey ("Ctrl", "a")
        pyautogui.press ("Delete")
        sleep(1.5)

        pyautogui.write(nome)
        sleep(2)
        pyautogui.press('f2')
        pyautogui.press('tab')
        sleep(2)
        pyautogui.press('menu')
        pyautogui.press('enter')
        sleep(2)

        pyautogui.press ("up")
        pyautogui.hotkey ("Ctrl", "a")
        pyautogui.press ("Delete")
        sleep(1.5)

        pyautogui.write (email)
        pyautogui.press ("space")
        pyautogui.hotkey ("ctrl", "a")
        pyautogui.hotkey ("ctrl", "z")
        pyautogui.write (complemento)
        sleep(3)

        pyautogui.press ("f2")
        pyautogui.press ("tab")
        pyautogui.press ("tab")
        sleep(2)
        pyautogui.press ("menu")
        pyautogui.press ("enter")
        sleep(2)
        pyautogui.hotkey ("Ctrl", "a")
        pyautogui.press ("Delete")
        sleep(1.5)

        pyautogui.write (cargo)
        pyautogui.press ("f2")

        sleep(2)
        pyautogui.hotkey ("Alt", "g")
        sleep(1.5)
        for i in range (6):
            pyautogui.press ("tab")
        pyautogui.press ("Enter")
        sleep(2)
        for i in range (5):
            pyautogui.press("down")
        pyautogui.press ("Enter")
        sleep(2)
        for i in range (5):
            pyautogui.press("down")
        pyautogui.press ("Enter")
        pyautogui.press ("Enter")
        pyautogui.press ("Enter")
        sleep(2)

        navegador.execute_script("window.open('about:blank', '_blank');")
        sleep(2)
        navegador.switch_to.window(navegador.window_handles[1])
        sleep(2)
        navegador.get('https://outlook.office.com/mail/')
        sleep(6)
        novoemail= navegador.find_element (By.XPATH , '//*[@id="innerRibbonContainer"]/div[1]/div/div/div/div[1]/div/div/span/button[1]/span').click()
        sleep(3)        
        destinatario = navegador.find_element (By.XPATH , '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[1]/div/div[5]/div/div').click()
        sleep(1)

        pyautogui.write (email)
        for i in range (3):
            pyautogui.press ("tab")
        pyautogui.write ("Assinatura Eletronica")
        sleep(2)
        pyautogui.press ("tab")
        pyautogui.write (mensagem)
        pyautogui.press ("enter")
        sleep(2)

        anexar = navegador.find_element (By.XPATH, '//*[@id="innerRibbonContainer"]/div[4]/div/div/div/div[1]').click()
        sleep(2)    
        pyautogui.press ("tab")
        pyautogui.press ("enter")
        sleep(5)
        pyautogui.write("Downloads")
        pyautogui.press("enter")
        sleep(2)
        for i in range (11):
            pyautogui.press ("tab")
        sleep(2)
        pyautogui.press("down")
        pyautogui.press("up")
        pyautogui.press ("enter")
        sleep(3)

        anexar = navegador.find_element (By.XPATH, '//*[@id="innerRibbonContainer"]/div[4]/div/div/div/div[1]').click()
        sleep(2)
        pyautogui.press ("tab")
        pyautogui.press ("enter")
        sleep(6)
        pyautogui.write ("Manual para Assinatura.pdf")
        pyautogui.press ("enter")
        sleep(6)

        enviaremail = navegador.find_element (By.XPATH , ' //*[@id="docking_InitVisiblePart_0"]/div/div[2]/div[1]/div/span/button[1]/span').click()
        print ("FIM DO PROGRAMA")    
 
root = Tk()
outlook_automation = OutlookAutomation()
interface = InterfaceGrafica(root,outlook_automation)
root.mainloop()

#terminal >> pip install mouseinfo >> python >> from mouseinfo import mouseInfo >> mouseInfo() >> Enter /apertar f6 na coordenada desejada ira gravar...