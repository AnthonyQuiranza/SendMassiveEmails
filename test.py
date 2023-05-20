import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
load_dotenv()

# HTML message
html = '''
    <html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EDUVIRTUAL</title>

<style type="text/css">
body {
	margin: 0;
}
body, table, td, p, a, li, blockquote {
	-webkit-text-size-adjust: none!important;
	font-family: sans-serif;
	font-style: normal;
	font-weight: 400;
}
button {
	width: 90%;
}

@media screen and (max-width:600px) {
/*styling for objects with screen size less than 600px; */
body, table, td, p, a, li, blockquote {
	-webkit-text-size-adjust: none!important;
	font-family: sans-serif;
}
table {
	/* All tables are 100% width */
	width: 100%;
}
.footer {
	/* Footer has 2 columns each of 48% width */
	height: auto !important;
	max-width: 48% !important;
	width: 48% !important;
}
table.responsiveImage {
	/* Container for images in catalog */
	height: auto !important;
	max-width: 30% !important;
	width: 30% !important;
}
table.responsiveContent {
	/* Content that accompanies the content in the catalog */
	height: auto !important;
	max-width: 66% !important;
	width: 66% !important;
}
.top {
	/* Each Columnar table in the header */
	height: auto !important;
	max-width: 48% !important;
	width: 48% !important;
}
.catalog {
	margin-left: 0%!important;
}
}

@media screen and (max-width:480px) {
/*styling for objects with screen size less than 480px; */
body, table, td, p, a, li, blockquote {
	-webkit-text-size-adjust: none!important;
	font-family: sans-serif;
}
table {
	/* All tables are 100% width */
	width: 100% !important;
	border-style: none !important;
}
.footer {
	/* Each footer column in this case should occupy 96% width  and 4% is allowed for email client padding*/
	height: auto !important;
	max-width: 96% !important;
	width: 96% !important;
}
.table.responsiveImage {
	/* Container for each image now specifying full width */
	height: auto !important;
	max-width: 96% !important;
	width: 96% !important;
}
.table.responsiveContent {
	/* Content in catalog  occupying full width of cell */
	height: auto !important;
	max-width: 96% !important;
	width: 96% !important;
}
.top {
	/* Header columns occupying full width */
	height: auto !important;
	max-width: 100% !important;
	width: 100% !important;
}
.catalog {
	margin-left: 0%!important;
}
button {
	width: 90%!important;
}
}
</style>
</head>
<body yahoo="yahoo">
<table width="100%"  cellspacing="0" cellpadding="0">
  <tbody>
    <tr>
      <td><table width="600"  align="center" cellpadding="0" cellspacing="0">
          <!-- Main Wrapper Table with initial width set to 60opx -->
          <tbody>
            <tr>
              <td><table bgcolor="#006633" class="top" width="98%"  align="left" cellpadding="0" cellspacing="0" style="padding:10px 10px 10px 10px;">
                  <!-- First header column with Logo -->
                  <tbody>
                    <tr>
                      <td style="font-size: 12px; color:#9D9D9C; text-align:center; font-family: sans-serif;">
						<img width="50%" src="https://eduvirtual.upec.edu.ec/assets/img/logo_white.png">
						</td>
                    </tr>
                  </tbody>
                </table>
                </td>
            </tr>
            <tr> 
              <!-- HTML Spacer row -->
              <td style="font-size: 0; line-height: 0;" height="20"><table width="96%" align="left"  cellpadding="0" cellspacing="0">
                  <tr>
                    <td style="font-size: 0; line-height: 0;" height="20">&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
            <tr> 
              <!-- HTML Spacer row -->
              <td style="font-size: 0; line-height: 0;" height="20"><table width="96%" align="left"  cellpadding="0" cellspacing="0">
                  <tr>
                    <td style="font-size: 0; line-height: 0;" height="20">&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
            <tr> 
              <!-- Introduction area -->
              <td><table width="96%"  align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <!-- row container for TITLE/EMAIL THEME -->
                    <td align="center" style="font-size: 32px; font-weight: 300; line-height: 2.5em; color: #9D9D9C; font-family: sans-serif;">¡Bienvenid@ a la UPEC!</td>
                  </tr>
                  <tr> 
                    <!-- row container for Tagline -->
                    <td align="center" style="font-size: 16px; font-weight:300; color: #9D9D9C; font-family: sans-serif;"><strong>EDUVIRTUAL</strong> te da la bienvenida </td>
                  </tr>
                  <tr>
                    <td style="font-size: 0; line-height: 0;" height="20"><table width="96%" align="left"  cellpadding="0" cellspacing="0">
                        <tr> 
                          <!-- HTML Spacer row -->
                          <td style="font-size: 0; line-height: 0;" height="20">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr> 
                    <!-- Row container for Intro/ Description -->
                    <td align="left" style="font-size: 14px; font-style: normal; font-weight: 100; color: #9D9D9C; line-height: 1.8; text-align:justify; padding:10px 20px 0px 20px; font-family: sans-serif;">
						¡Bienvenido a la Universidad Politécnica Estatal del Carchi! Desde la Unidad de Tecnología Educativa <strong> EDUVIRTUAL</strong> Estamos emocionados de recibir a nuestros nuevos estudiantes de Nivelación. Prepárate para explorar, aprender y crecer juntos en esta emocionante etapa de tus estudios ¡Te deseamos mucho éxito en tu camino académico! <br>
						¡Hemos preparado algo para ti! <br> Los inicios suelen ser un poco difíciles, es por esto que creamos un video introductorio para que puedas aprender a usar tu aula virtual eduvirtual.upec.edu.ec. Al hacer clic en el siguiente botón, encontrarás nuestro video:
					  <a href="#" style="text-decoration:none">
                                      <p style="background-color:#006633; text-align:center; padding: 10px 10px 10px 10px; margin: 10px 10px 10px 10px;color: #FFFFFF;   font-family: sans-serif; ">Ver video introductorio</p>
                                      </a></td>
					  
                  </tr>
                </table></td>
            </tr>
            
           
          </tbody>
        </table></td>
    </tr>
  </tbody>
</table>
</body>
</html>
    '''

# get environment variables
user=os.getenv("USER")
password=os.getenv("PASSWORD")
# get information about mail
emailmessage= MIMEMultipart()
emailmessage['From'] = f'{user}'
emailmessage['Subject'] = 'Introducción a la plataforma Eduvirtual'
emailmessage.attach(MIMEText(html, 'html'))
mailserver = smtplib.SMTP('smtp.office365.com',587)
mailserver.ehlo()
mailserver.starttls()
mailserver.ehlo()
#Login to the server
mailserver.login(f'{user}', f'{password}')
#Reading data
data = pd.read_excel("Correos prueba.xlsx")
emails=data.columns[0]
#Sending emails
for email in range(len(data[emails])):
    print()
    mailserver.sendmail(f'{user}', data[emails][email], emailmessage.as_string())
    print(f'Correo enviado correctamente a {data[emails][email]}')
    
#Close connection
mailserver.quit()
