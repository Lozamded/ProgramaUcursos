# main.py
from appJar import gui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from screeninfo import get_monitors
import time, random, os
from pdb import set_trace as bp


class MenuApp():

	def __init__(self, browser):
		self.browser = browser

	def Subir(self,btnName):
			print("Subir HTML")
			App3 = SubirApp(self.browser)
			self.app.stop()
			App3.Start3()

	def Arbol(self,btnName):
			print("Crear Arbol")
			App4 = ArbolApp(self.browser)
			self.app.stop()
			App4.Start4()

	def Prepare2(self, app):
		app.setBg('#B8B3A4')
		app.setResizable(canResize=False)
		m = get_monitors()
		app.setLocation(m[0].width*(0.6), m[0].height*(0.2))
		app.setPadX(5)

		app.addImage("title", 'ucursos3.gif', 0, 0, 2)
		app.zoomImage("title", -2)

		app.addNamedButton("Subir HTML","Subir",self.Subir, 2, 0)
		app.setButtonSticky('Subir', 'right')

		app.addNamedButton("Ingresar Arbol","Arbol",self.Arbol, 3, 0)
		app.setButtonSticky('Arbol', 'right')

		return app

	# Build and Start your application
	def Start2(self):
		# Creates a UI
		print("Menu")
		app = gui('U-cursos.cl')
		app = self.Prepare2(app)
		self.app = app
		app.go()


class SubirApp():
	
	def __init__(self, browser):
		self.browser = browser

	def SubmitHTML(self,btnName):
		self.app.setLabel("HTMLwebpage",self.browser.current_url)	
		self.browser.get(self.browser.current_url)
		
	def SaveZipDir(self,btnName):
		dirZip = self.app.directoryBox(title="Buscar .zip en...", dirName=None)
		self.app.setLabel("ZipDir",dirZip)
		
	def StartUpload(self,btnName):
		if self.app.getLabel('HTMLwebpage') != "" and self.app.getLabel("ZipDir") != "":
			ZipDir = self.app.getLabel("ZipDir")
			ZipFilesPrev = os.listdir(self.app.getLabel("ZipDir"))
			for file in ZipFilesPrev:
				if file.endswith(".zip"):
					time.sleep(1)
					boton = self.browser.find_element_by_class_name('boton').click()
					time.sleep(1)
					self.browser.find_element_by_name('h[titulo]').send_keys(file.replace('.zip',''))
					self.browser.find_element_by_name("archivo[]").send_keys(ZipDir+"/"+file)
					self.browser.find_element_by_xpath("/html/body/div[@id='derecha']/div[@id='body']/form[1]/input[4]").click()
					#time.sleep(random.uniform(3,5))
					try:
						element = WebDriverWait(self.browser, 100).until(EC.presence_of_element_located((By.ID, "mexito")))
					finally:
						pass
			
	# Build the GUI
	def Prepare3(self, app):#Interfaz de subida
		# Form GUI
		app.setBg('#B8B3A4')
		app.setResizable(canResize=False)
		m = get_monitors()
		app.setLocation(m[0].width*(0.55),m[0].height*(0.2))

		#app.setGuiPadding(10, 10)
		app.setPadX(5)
		#app.setIcon('ind.ico')
		app.addImage("title",'ucursos3.gif',0,0,2) 
		app.zoomImage("title",-2)

		# HTML site
		app.addLabel("HTMLwebpageTag", "Pagina de HTML", 1, 0)		
		app.setLabelAlign("HTMLwebpageTag","left")		
		app.addLabel("HTMLwebpage","", 1, 1)
		app.setLabelAlign("HTMLwebpage","left")
		app.setLabelWidth("HTMLwebpage",40)
		app.setLabelBg('HTMLwebpage','white')
		app.addNamedButton("Guardar HTML",'SaveHTML',self.SubmitHTML,1,2)	#llama la funcion para subir el html	
		app.setButtonSticky('SaveHTML','right')

		# ZIP DIR
		app.addLabel("ZipDirTag", "Carpeta .Zip", 2, 0)		
		app.setLabelAlign("ZipDirTag","left")		
		app.addLabel("ZipDir","", 2, 1)
		app.setLabelAlign("ZipDir","left")
		app.setLabelWidth("ZipDir",30)
		app.setLabelBg('ZipDir','white')
		app.addNamedButton("Subir Zip",'SaveZip',self.SaveZipDir,2,2)	#Llama a la funcion para subir los .zip
		app.setButtonSticky('SaveZip','right')
		
		
		app.addNamedButton("Iniciar",'StartUpload',self.StartUpload,3,2)	
		app.setButtonSticky('StartUpload','right')
		
		return app

	# Build and Start your application
	def Start3(self):
		# Creates a UI
		app = gui('U-cursos.cl')
		app = self.Prepare3(app) #LLama la interfaz de subida
		self.app = app
		app.go()


class ArbolApp():

	def __init__(self, browser):
		self.browser = browser
		
	def Prepare4(self,app):  #Interfaz de subida
		# Form GUI
		print("Algo")
		app.setBg('#B8B3A4')
		app.setResizable(canResize=False)
		m = get_monitors()
		app.setLocation(m[0].width*(0.55), m[0].height*(0.2))

		##app.setGuiPadding(10, 10)
		app.setPadX(5)
		##app.setIcon('ind.ico')
		app.addImage("title", 'ucursos3.gif', 0, 0, 2)
		app.zoomImage("title", -2)

		return app

	def Start4(self):
		# Creates a UI
		app = gui('U-cursos.cl')
		app = self.Prepare4(app) #LLama la interfaz de subida
		self.app = app
		app.go()
	

class loginApp():	
    # Build the GUI
    def Prepare(self,app):
		# Form GUI
		app.setBg('#B8B3A4')
		app.setResizable(canResize=False)
		m = get_monitors()
		app.setLocation(m[0].width*(0.6),m[0].height*(0.2))
		#app.setGuiPadding(10, 10)
		app.setPadX(5)
		#app.setIcon('ind.ico')
		app.addImage("title",'ucursos3.gif',0,0,2) 
		app.zoomImage("title",-2)
		#Texto Usuario 		
		app.addLabel("userLab", "Usuario", 1, 0)		
		app.setLabelAlign("userLab","left")		
		#Texto Password
		app.addLabel("passLab", "Password", 2, 0)
		app.setLabelAlign("passLab","left")
		#Ingresables
		app.addEntry("username", 1, 1)
		app.setEntryAlign("username","left")
		app.setEntryWidth("username",30)			
		app.addSecretEntry("password", 2, 1)
		app.setEntryAlign("password","left")
		app.setEntryWidth("password",30)	

		#Driver
		app.addLabel("DriverDirTag", "Chromedriver", 3, 0)		
		app.setLabelAlign("DriverDirTag","left")		
		app.addLabel("DriverDir","", 3, 1)
		app.setLabelAlign("DriverDir","left")
		app.setLabelWidth("DriverDir",30)
		app.setLabelBg('DriverDir','white')
		app.addNamedButton("Buscar",'Search',self.Driver,3,2)
		app.setButtonSticky("Search","right")

		#botones
		app.addNamedButton("Ingresar","Ingresar",self.Submit,4,1)
		app.setButtonSticky('Ingresar','right')
		app.setFocus("username")	
		app.enableEnter(self.Submit)
		
		return app
        


    # Build and Start your application
    def Start(self):
        # Creates a UI
        app = gui('Ingresar a http://u-cursos.cl')
        app = self.Prepare(app)
        self.app = app
        app.go()

    # Define method that is executed when the user clicks on the submit buttons
    # of the form

    def Driver(self, btnName):
		dirDriver = self.app.openBox(title="Buscar chromedriver en...", dirName="C:User/desktop",asFile = False)
		print("Chromedriver: ")
		print(dirDriver)
		self.app.setLabel("DriverDir",dirDriver)
		return dirDriver

    def Submit(self,btnName): #funcion del boton de subida
		email = self.app.getEntry("username")
		password = self.app.getEntry("password")
		dirDriver = self.app.getLabel("DriverDir")
		print("Dir driver es " + dirDriver)
		
		if email != "" and password != "":
		
			chrome_options = Options()
			chrome_options.add_argument("--window-size=800,700")
			if(dirDriver == ""): #si  el usuario no ha ingresafo nada en el Driver
				browser = webdriver.Chrome(chrome_options=chrome_options)
			else:
				browser = webdriver.Chrome(executable_path=dirDriver, chrome_options=chrome_options)

			browser.get("https://www.u-cursos.cl/")

			emailElement = browser.find_element_by_name("username")
			emailElement.send_keys(email)
			passElement = browser.find_element_by_name("password")
			passElement.send_keys(password)
			passElement.submit()
			time.sleep(random.uniform(3.5,6.9))
			
			#os.system('cls')
			#print "[+] Success! Logged In, Bot Starting!"
			try:
				browser.find_element_by_class_name('mensaje')
				self.app.errorBox("Error de login","Su Usuario o password son incorrectos")
				#print("cerrar el navegador")
				browser.close()
			except NoSuchElementException:
				# estamos adentro de ucursos
				print("Desplegar Menu")
				App2 = MenuApp(browser)
				self.app.stop()
				App2.Start2()
		else:
			self.app.errorBox("Error de login", "Usuario o password invalido")
	

if __name__ == '__main__':
    App = loginApp()
    App.Start()
	
