Para usar esta plantilla con la autenticación de Windows Azure, consulte http://go.microsoft.com/fwlink/?LinkID=267940.

De lo contrario, para usar esta plantilla con la autenticación de Windows, consulte las instrucciones a continuación:

Hospedaje en IIS Express:
1. Haga clic en el proyecto en el Explorador de soluciones para seleccionar el proyecto.
2. Si el panel Propiedades no se abre, ábralo (F4).
3. En el recuadro Propiedades del proyecto:
   a) Establezca "Autenticación anónima" en "Deshabilitada".
   b) Establezca "Autenticación de Windows" en "Habilitada":

Hospedaje en IIS 7 o posterior:
1. Abra el Administrador de IIS y navegue a su sitio web.
2. En Vista Características, haga doble clic en Autenticación.
3. En la página Autenticación, seleccione Autenticación de Windows. Si la opción Autenticación de Windows
   no aparece, asegúrese de que la característica Autenticación de Windows
   esté instalada en el servidor.

      Para habilitar la autenticación de Windows en Windows:
      a) En el Panel de control abra "Programas y características".
      b) Seleccione "Activar o desactivar las características de Windows".
      c) Navegue hasta Internet Information Services > Servicios World Wide Web > Seguridad
         y asegúrese de que la casilla del nodo Autenticación de Windows está activada.

      Para habilitar la autenticación de Windows en Windows Server:
      a) En Administrador de servidores, seleccione Servidor web (IIS) y haga clic en Agregar servicios de rol.
      b) Navegue a Servidor web > Seguridad
         y asegúrese de que la casilla del nodo Autenticación de Windows está activada.

4. En el panel Acciones, haga clic en Habilitar para utilizar la autenticación de Windows.
5. En la página Autenticación, seleccione Autenticación anónima.
6. En el panel Acciones, haga clic en Deshabilitar para deshabilitar la autenticación anónima.