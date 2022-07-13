import pandas as pd
from xml.etree.ElementTree import Element, SubElement, Comment, tostring
from xml.etree import ElementTree
import time
from xml.dom import minidom
from os import listdir
from os.path import isfile, join
from zipfile import ZipFile

def cross(path,namefile):
    logging=""
    # departments
    deps = pd.read_excel("./Data/Deps.xlsx")
    ndeps = []
    for i in deps["BAS_DEPARTAMENTOS"].tolist():
        ndeps.append(str(i).replace(" ","").upper())
    deps["BAS_DEPARTAMENTOS"]=ndeps
    def getdep(name):
        name = name.replace(" ","").upper().replace("D.C","").replace("Á",'A').replace("Ó","O").replace("Í","I").replace(".","").replace(",","").replace("Ò","O")
        if("VALLEDELCAUCA" == name):
            name = "VALLE"
        if("NARIÑO" == name):
            name = "NARINO"
        if("LAGUAJIRA" == name):
            name = "GUAJIRA"
        return str(deps[deps["BAS_DEPARTAMENTOS"] == name]["cod"].tolist()[0])
    # municipios
    mun = pd.read_excel("./Data/Mun.xlsx")
    nmun = []
    for i in mun["NOMBRE_CIUDAD"].tolist():
        nmun.append(str(i).replace(" ","").upper())
    mun["NOMBRE_CIUDAD"]=nmun
    def getmuni(name,dep):
        name =name.replace(" ","").upper().replace("D.C","").replace("Á",'A').replace("Ó","O").replace("Í","I").replace("Ñ","N").replace("É","E").replace("Ú","U").replace('Ü',"U").replace(".","").replace(",","")
        if(name == "CARTAGENA"):
            name="CARTAGENADEINDIAS"
        if(name == "ELCARMENDEVIBORAL"):
            name="CARMENDEVIBORAL"
        if(name == "ISTMINA"):
            name="ITSMINIA"
        if(name == "INIRIDA"):
            name="INHIRIDA"
        if(name == "SANJOSEDELFRAGUA"):
            name="SANJOSEDELAFRAGUA"
        cities = mun["NOMBRE_CIUDAD"].tolist()
        codes = mun["CODIGO_CIUDAD"].tolist()
        findnames = ""
        for i in range(len(cities)):
            if (name==cities[i] and str(int(codes[i])).startswith(str(dep))):
                return codes[i]
        return str(mun[mun["NOMBRE_CIUDAD"] == name]["CODIGO_CIUDAD"].tolist()[0])
    Cargue = pd.read_excel(path)
    planes= pd.read_excel("./Data-plan/BasePlanes.xlsx")
    empr= pd.read_excel("./Data-plan/Base.xlsx",'emprendedor')
    inventory = pd.read_excel("./Data-plan/Inventarios.xlsx")
    ciiu = pd.read_excel("./Data-plan/CIUUFinal.xlsx")
    alldata = pd.merge(Cargue,planes,how="left",left_on="ID Plan de Negocio",right_on="id_plan")
    alldata = pd.merge(alldata,empr,how="left",left_on="ID Plan de Negocio",right_on="id_plan")
    alldata = pd.merge(alldata,inventory,how="left",left_on="ID Plan de Negocio",right_on="ID Plan de Negocio")
    alldata = pd.merge(alldata,ciiu,how="left",left_on="ciiu",right_on="Clasecod")
    datadict = alldata.to_dict()
    muncode = []
    depcode =[]
    print("Crossing...")
    logging = logging + "Crossing... \n"
    for i in range(len(datadict["municipio"])):
        try:
            dcode = getdep(datadict["departamento"][i])
            depcode.append(dcode)
        except:
            depcode.append("")
        try:
            muni = getmuni(datadict["municipio"][i],dcode)
            muncode.append(muni)
        except:
            muncode.append("")
    alldata["Muncode"]=muncode
    alldata["Depcode"]=depcode
    alldata.to_excel("./Data-plan/"+namefile+".xlsx")
    logging+= "export to ./Data-plan/"+namefile+".xlsx \n" 
    return logging
def export(exportpath,filepath,n):
    logging =""
    bienes = [
    '1 Equipo industrial',
    '2 Equipo construcción',
    '3 Equipo oficina',
    '4 Equipo agrícola',
    '5 Otro equipo',
    '6 Productos agrícolas',
    '7 Inventarios',
    '8 Vehículos',
    '9 Cuentas por cobrar',
    '10 Bienes por adhesión',
    '15 Acciones o participaciones en el capital']
    bcom = {
        "comercial":"1",
        "consumo":"2",
        "ambos":"3"
    }
    def prettify(elem):
        """Return a pretty-printed XML string for the Element.
        """
        rough_string = ElementTree.tostring(elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")

    def addValueSub(root,name,val):
        teme = SubElement(root,name)
        teme.text = val
    from datetime import datetime
    def getnamehour():
        today = datetime.now()
        st = "6_899999034_"
        return st + str(today.year) + getcomplete(str(today.month)) +getcomplete(str(today.day))+getcomplete(str(today.hour))+getcomplete(str(today.minute))+getcomplete(str(today.second))

    def getcomplete(s):
        return s if len(s)>=2 else "0"+s
    def raplaces(text):
        chars=  ["#","$","ª","º","®","º","¼","½","Á","É","Ë","Í","Ñ","Ó","×","Ú","-","'",'"','"',"™" ,"≤", "%","°",       "-","_","/","“","”","³","’"]
        carrep =["No.","","","","","","1/4","1/2","A","E","E","I","N","O","x","U","","","","","",""      , "" ," grados ","" , "", "", "", "", "", ""]
        for i in range(len(chars)):
            text = text.replace(chars[i],carrep[i])
        return text
    def escape_html(text):
        """Escape &, <, > as well as single and double quotes for HTML."""
        text = text.replace('&',"" ).replace('"',"").replace('&quot;',"").replace("&amp;","").replace("&lt;","<").replace("&gt;",">")
        return raplaces(text)
    def getinfoxml(top,rdi):
        gcl=SubElement(top, 'gcl')
        #----------------------------Deudor
        ddor=SubElement(gcl, 'ddor')
        addValueSub(ddor,"cci",'6') #Cambiar por los nuevos campos "foelec", "ref", "canpor". Preguntar a Jessica por estos campos
        addValueSub(ddor,"ni",str(rdi["nit"]))
        addValueSub(ddor,"dv",str(rdi["digito_verificacion"]))
        addValueSub(ddor,"rs",str(rdi["razon_social"]))
        addValueSub(ddor,"pais",'CO')
        addValueSub(ddor,"dpto",str(rdi["Depcode"]))
        muncode = str(rdi["Muncode"])
        addValueSub(ddor,"mun", muncode if len(muncode)>4 else "0"+muncode)
        addValueSub(ddor,"dir",str(rdi["Direccion de la Empresa\n(Registro Confecámaras)"]))
        addValueSub(ddor,"email",str(rdi["email final"]))
        addValueSub(ddor,"tel",str(int(rdi["tel final"])))
        addValueSub(ddor,"cel",str(int(rdi["tel final"])))
        addValueSub(ddor,"tdc",'0')
        addValueSub(ddor,"ins",'false')
        addValueSub(ddor,"emtam",'1')
        addValueSub(ddor,"pf",'true' if str(rdi["Género del Emprendedor\n(Seleccionar F ó M)"])=="F" else "false" )
        #sector no esta igual
        sec = SubElement(ddor, 'sec')
        #print(str(rdi["Sector"]))
        addValueSub(sec,"cod",str(int(rdi["sec"])))
        addValueSub(ddor,"tddor",'g')

        #----------------------------Acreedoor
        acdor=SubElement(gcl, 'acdor')
        addValueSub(acdor,"cci",'6')
        addValueSub(acdor,"ni",'899999034')
        addValueSub(acdor,"dv",'1')
        addValueSub(acdor,"rs",'Servicio Nacional de Aprendizaje SENA')
        addValueSub(acdor,"pais",'CO')
        addValueSub(acdor,"dpto",'11')
        addValueSub(acdor,"mun",'11001')
        addValueSub(acdor,"dir",'Calle 57 No. 8-69')
        #no info
        addValueSub(acdor,"email",'garantias@sena.edu.co')
        addValueSub(acdor,"tel",'5492080')
        addValueSub(acdor,"cel",'3204853139')

        addValueSub(acdor,"ppal",'true')
        addValueSub(acdor,"ppar",'100.0')

        #------------------------------- Bienes

        addValueSub(gcl,"descbien",escape_html(str(rdi["Descripcion de Bienes\n(Inventario)"])))
        addValueSub(gcl,"prad",'true')
        addValueSub(gcl,"monto",str(rdi["Valor Total de los Bienes\n(Inventario)"]))
        #fecha
        addValueSub(gcl,"vdef",'false')
        #addValueSub(gcl,"ffin",'0')

        addValueSub(gcl,"ctg",'1')
        addValueSub(gcl,"cm",'COP')
        addValueSub(gcl,"cbu",bcom[str(rdi["BAS_BIENES_USO\n(Comercial, Consumo, Ambos)"]).lower()])
        for b in bienes:
            if(rdi[b] and str(b.split()[0])!="15"):
                cbien = SubElement(gcl, 'cbien')
                addValueSub(cbien,"cod",str(b.split()[0]))
    print("running...")
    logging = logging + "Running... \n"
    alldata = pd.read_excel(filepath)
    alldata = alldata[pd.notna(alldata["ID Plan de Negocio"])]
    rd = alldata.to_dict(orient='records')
    for i in range(0,len(rd),n):
        rds =rd[i:i+n]
        top = Element('garantias')
        op= SubElement(top, 'op')
        addValueSub(op,"t","I") #CAMBIAR LA 'I' POR 'C'
        addValueSub(op,"tg",str(len(rds))) #Agregar un tab
        n=0
        for rdi in rds:
            try:
                getinfoxml(top,rdi)
                n+=1
            except Exception as e:
                print("error on " + str(rdi['ID Plan de Negocio']))
                print(str(e))
        print("File "+str(i))
        logging = logging + "File "+str(i) +"\n"
        xmltext = prettify(top)
        # print(xmltext)
        name = exportpath+getnamehour()+".xml"
        outF = open(name, "w",encoding='utf8')
        outF.writelines(xmltext)
        print(name)
        logging = logging + name + " \n"
        print(n)
        logging = logging + str(n) + " \n"
        outF.close()
        time.sleep(11)
    return logging
import wx

class OtherFrame(wx.Frame):
    """
    Class used for creating frames other than the main one
    """
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title,pos=(60,60))
        self.panel = wx.Panel(self)
        self.my_sizer = wx.BoxSizer(wx.VERTICAL)
    def print_on_frame(self,text):
        textelement = wx.StaticText(self.panel)
        textelement.SetLabel(text)
        self.my_sizer.Add(textelement, 0, wx.ALL | wx.EXPAND, 5)
        self.panel.SetSizer(self.my_sizer)
        self.Show()

class MyFrame(wx.Frame):    
    def __init__(self):
        super().__init__(parent=None, title='Generador de XMl')
        panel = wx.Panel(self)        
        my_sizer = wx.BoxSizer(wx.VERTICAL)        
        self.text_ctrl = wx.TextCtrl(panel)
        my_sizer.Add(self.text_ctrl, 0, wx.ALL | wx.EXPAND, 5)        
        my_btn = wx.Button(panel, label='Selecionar Archivo de IDs')
        my_btn.Bind(wx.EVT_BUTTON, self.on_press)
        my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)
        self.name_ctrl = wx.TextCtrl(panel)
        self.name_ctrl.SetValue("Nombre archivo exportar")
        my_sizer.Add(self.name_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        self.n_ctrl = wx.TextCtrl(panel)
        self.n_ctrl.SetValue("Numero de planes por XML")
        my_sizer.Add(self.n_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        btn_process = wx.Button(panel, label='Process')
        btn_process.Bind(wx.EVT_BUTTON, self.process)
        my_sizer.Add(btn_process, 0, wx.ALL | wx.CENTER, 5)
        # set sizeer
        panel.SetSizer(my_sizer)
        
        # show interface
        self.Show()
    def on_press(self,event):
        # Create open file dialog
        openFileDialog = wx.FileDialog(frame, "Open", "", "", 
            "*", 
            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        self.text_ctrl.SetValue (openFileDialog.GetPath())
    def process(self,event):
        path = self.text_ctrl.GetValue()
        namefile=self.name_ctrl.GetValue()
        exportpath = "./export/"
        filepath = "./Data-plan/"+namefile+".xlsx"
        n = int(self.n_ctrl.GetValue())
        logs=""
        try:
            logs += cross(path,namefile)
            logs+= export(exportpath,filepath,n)
        except Exception as e:
            log = str(e)
            print(e)
        print("LOG",logs)
        self.frame = OtherFrame(title="logging")
        self.frame.print_on_frame(logs)
        
if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()