import uno
import unohelper
from org.user.miprimerafuncion import XNumeroAMoneda

ID_EXTENSION = 'org.user.miprimerafuncion'
SERVICE = ('com.sun.star.sheet.AddIn',)

Menosde21 = ['Cero','Un','Dos','Tres','Cuatro','Cinco','Seis','Siete','Ocho','Nueve','Diez',
                  'Once','Doce','Trece','Catorce','Quince','Dieciséis','Diecisiete',
                  'Dieciocho','Diecinueve','Veinte']
                  
Menosde500 = {2:'Veinti',3:'Treinta',4:'Cuarenta',5:'Cincuenta',6:'Sesenta',7:'Setenta',8:'Ochenta',9:'Noventa',100:'Cien',500:'Quinientos '}
            
class NumeroAMoneda(unohelper.Base, XNumeroAMoneda):

    def __init__(self, ctx):
        self.ctx = ctx

    def CantAMoneda(self, value):
        #numero = value.replace(',','.') #reemplazo las comas por puntos por si acaso
        f_numero = float(value)
        entero = int(f_numero//1)
        decimal = int(100*round(f_numero%1,2))
        print (entero)
        print (decimal)
        cadena_importe = self.PasarATexto(entero)        
        cadena_importe += " EUROS "
        if decimal>0:
            cadena_importe += "CON "
            cadena_importe += self.PasarATexto(decimal)
            cadena_importe += " CÉNTIMO"
            if decimal>1:
                cadena_importe+="S"
        cadena_importe = cadena_importe.upper()
        print (cadena_importe)
        return cadena_importe
	
    def PasarATexto(self,numero):
        cadena_numero=""
        if numero>0 and numero<21:  #intevalo 0-20
            cadena_numero = Menosde21[int(numero)]
        elif numero>20 and numero<100: #intervalo 21-99
            cadena_numero = Menosde500[int(numero/10)]
            if (numero >29 and numero%10>0): #subintervalo 30,40,50,60,70,80,90
                cadena_numero+= " y " 
            cadena_numero += self.PasarATexto (int(numero%10))
        elif numero==100: #intervalo 100
            cadena_numero = Menosde500[numero]
        elif (numero>100 and numero<500) or (numero>599 and numero<1000): #intervalo 101-499,600-999
            if numero>199: #subintervalo 200-499,600-999
                cadena_numero=self.PasarATexto(int(numero/100)) + "cientos "
            else:
                cadena_numero = "Ciento "
            cadena_numero += self.PasarATexto(int(numero%100))
        elif numero>499 and numero<600: #intervalo 500-599
            cadena_numero = Menosde500[int(numero-numero%100)] + self.PasarATexto(int(numero%100))
        elif numero>999 and numero<1000000: #intervalo 1000-999999
            if numero>1999: #subintervalo 2000-999999
                cadena_numero = self.PasarATexto(int(numero/1000))
            cadena_numero+=" Mil "
            cadena_numero+=self.PasarATexto(int(numero%1000))
        elif numero>999999 and numero<1000000000: #intervalo 1000000-999999999
            cadena_numero = self.PasarATexto(int(numero/1000000))
            if numero>1999999: #subintervalo 2000000-999999999
                cadena_numero+=" Millones "
            else:
                cadena_numero+= " Millón "
            cadena_numero+=self.PasarATexto(int(numero%1000000))
        elif numero>999999999 and numero<1000000000000: #intervalo 1000000000-999999999999
            if numero>1999999999: #subintervalo 2000000000-999999999999
                cadena_numero = self.PasarATexto(int(numero/1000000000))
            cadena_numero+=" Mil "
            cadena_numero+=self.PasarATexto(int(numero%1000000000))                
        elif numero>999999999999 and numero<1000000000000000: #intervalo 1000000000000-999999999999999
            cadena_numero = self.PasarATexto(int(numero/1000000000000))
            if numero>1999999999999: #subintervalo 2000000000000-999999999999999
                cadena_numero+=" Billones "
            else:
                cadena_numero+=" Billón "
            cadena_numero+=self.PasarATexto(int(numero%1000000000000))                      
        #ultimos detalles
        #quitamos espacios al comienzo y al final
        cadena_numero = cadena_numero.strip()
        #cambiamos las centenas de 7 y 9
        cadena_numero = cadena_numero.replace ('Nuevecientos','Novecientos')
        cadena_numero = cadena_numero.replace ('Sietecientos','Setecientos')
        return cadena_numero
        	
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(NumeroAMoneda, ID_EXTENSION, SERVICE)


if __name__ == "__main__":
	Dato = NumeroAMoneda(XNumeroAMoneda)
	Dato.CantAMoneda("125451.01")
