from flask_wtf import Form
from wtforms import StringField, SubmitField, SelectField, DateField, BooleanField, FloatField, RadioField
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms.validators import DataRequired, Email
from flask_wtf.file import FileField

estados = ["Aguascalientes", "Baja California", "Baja California Sur", "Campeche",
           "Chiapas", "Chihuahua", "Ciudad de México", "Coahuila", "Colima", "Durango",
           "Estado de México", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco",
           "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", "Puebla",
           "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", "Sonora",
           "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"]

class Hardware(Form):
    #Vendedor
    nombre_vendedor = StringField('Nombre', validators=[DataRequired("Este campo es necesario")])
    cargo_vendedor = StringField('Cargo', validators=[DataRequired("Este campo es necesario")])
    mail_vendedor = StringField('Mail', validators=[DataRequired("Este campo es necesario"), Email("Escribe un email válido")])
    telefono_vendedor = StringField('Teléfono', validators=[DataRequired("Este campo es necesario")])
    #Cliente
    razon_social = StringField('Razón Social', validators=[DataRequired("Este campo es necesario")])
    cliente_corto = StringField('Nombre Corto', validators=[DataRequired("Este campo es necesario")])
    contacto_cliente = StringField('Contacto Con El Cliente', validators=[DataRequired("Este campo es necesario")])
    titulo_cliente = StringField('Titulo del Cliente', validators=[DataRequired("Este campo es necesario")])
    #Direccion
    estado_cliente = SelectField("Estado", choices = [(estado, estado) for estado in estados])
    ciudad_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    colonia_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    calle_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    postalCliente = StringField(validators=[DataRequired("Este campo es necesario")])
    #Contrato
    titulo_contrato = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_propuesta = StringField(validators=[DataRequired("Este campo es necesario")])
    years_garantia = StringField(validators=[DataRequired("Este campo es necesario")])
    precio_numero = FloatField('Precio', validators=[DataRequired("Escribe solo números")])
    #Archivo de configuracion
    config = FileField("Archivo de configuracion", validators=[FileRequired('Este campo es necesario'), FileAllowed(['txt'], 'Solo archivos txt')])
    fecha_contrato = DateField(validators=[DataRequired("Este campo es necesario")])
    fecha_vigencia = DateField(validators=[DataRequired("Este campo es necesario")])
    #Correo
    to_addr = StringField()
    submit = SubmitField('Enviar')

class Software(Form):
    #Vendedor
    nombre_vendedor = StringField('Nombre', validators=[DataRequired("Este campo es necesario")])
    cargo_vendedor = StringField('Cargo', validators=[DataRequired("Este campo es necesario")])
    mail_vendedor = StringField('Mail', validators=[DataRequired("Este campo es necesario"), Email("Escribe un email válido")])
    telefono_vendedor = StringField('Teléfono', validators=[DataRequired("Este campo es necesario")])
    #Cliente
    razon_social = StringField('Razón Social', validators=[DataRequired("Este campo es necesario")])
    cliente_corto = StringField('Nombre Corto', validators=[DataRequired("Este campo es necesario")])
    contacto_cliente = StringField('Contacto Con El Cliente', validators=[DataRequired("Este campo es necesario")])
    titulo_cliente = StringField('Titulo del Cliente', validators=[DataRequired("Este campo es necesario")])
    #Direccion
    estado_cliente = SelectField("Estado", choices = [(estado, estado) for estado in estados])
    ciudad_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    colonia_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    calle_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    postalCliente = StringField(validators=[DataRequired("Este campo es necesario")])
    #Contrato
    titulo_contrato = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_propuesta = StringField(validators=[DataRequired("Este campo es necesario")])
    precio_numero = FloatField('Precio', validators=[DataRequired("Escribe solo números")])
    #Archivo de configuracion
    config = FileField(validators=[FileAllowed(['docx'], 'Solo archivos de word (.docx)')])
    fecha_contrato = DateField(validators=[DataRequired("Este campo es necesario")])
    fecha_vigencia = DateField(validators=[DataRequired("Este campo es necesario")])
    #Correo
    to_addr = StringField()
    submit = SubmitField('Enviar')

class Servicios(Form):
    #Vendedor
    nombre_vendedor = StringField('Nombre', validators=[DataRequired("Este campo es necesario")])
    cargo_vendedor = StringField('Cargo', validators=[DataRequired("Este campo es necesario")])
    mail_vendedor = StringField('Mail', validators=[DataRequired("Este campo es necesario"), Email("Escribe un email válido")])
    telefono_vendedor = StringField('Teléfono', validators=[DataRequired("Este campo es necesario")])
    #Cliente
    razon_social = StringField('Razón Social', validators=[DataRequired("Este campo es necesario")])
    cliente_corto = StringField('Nombre Corto', validators=[DataRequired("Este campo es necesario")])
    contacto_cliente = StringField('Contacto Con El Cliente', validators=[DataRequired("Este campo es necesario")])
    titulo_cliente = StringField('Titulo del Cliente', validators=[DataRequired("Este campo es necesario")])
    #Direccion
    estado_cliente = SelectField("Estado", choices = [(estado, estado) for estado in estados])
    ciudad_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    colonia_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    calle_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_cliente = StringField(validators=[DataRequired("Este campo es necesario")])
    postalCliente = StringField(validators=[DataRequired("Este campo es necesario")])
    #Contrato
    titulo_contrato = StringField(validators=[DataRequired("Este campo es necesario")])
    numero_propuesta = StringField(validators=[DataRequired("Este campo es necesario")])
    years_garantia = StringField(validators=[DataRequired("Este campo es necesario")])
    precio_numero = FloatField('Precio', validators=[DataRequired("Escribe solo números")])
    OTC = RadioField('Forma de Pago', choices=[('OTC', 'Todo OTC'),('Mensual', 'Combinado (OTC - Mensual)')])
    #Archivo de configuracion
    config = FileField("Archivo de configuracion", validators=[FileRequired('Este campo es necesario'), FileAllowed(['txt'], 'Solo archivos txt')])
    fecha_contrato = DateField(validators=[DataRequired("Este campo es necesario")])
    fecha_vigencia = DateField(validators=[DataRequired("Este campo es necesario")])
    #Correo
    to_addr = StringField()
    submit = SubmitField('Enviar')

class Checkboxes(Form):
     serv1 = BooleanField("Incluir servicios de LAB services", default="checked")
     serv2 = BooleanField("Incluir servicios profesionales de IBM", default="checked")
     serv3 = BooleanField("Incluir servicios de Mantenimiento", default="checked")
     config = FileField("CHIS", validators=[FileRequired('Este campo es necesario'), FileAllowed(['docx'], 'Solo archivos txt')])
     submit = SubmitField('Enviar')
