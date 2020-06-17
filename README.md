# Creo-open-DRW
Allows opening a corresponding *.drw file of a *.prt or *.asm file in the cad program <a href="https://www.ptc.com/en/products/creo/parametric">Creo Parametrics</a>.<br/>
<b>Notice:</b> This is just a test version witch is not pretending to be super fast. It is just a <i>free</i> solution.
## Setup
Follow the steps provided at <a href="https://creocustomization.com/how-to-install-and-register-creo-vb-api-toolkit-component-in-creo-parametric/">this page</a>.<br/>
Just in case it is already down here a <a href="Setup/Creo%20Setup.htm">backup</a>.<br/><br/>
I have provided a sample setup (powershell) script <a href="Setup/Setup.ps1">here</a>.
## Make it work
Start Creo Parametrics with a batch (etc) file and start the application at the same time. At the moment it will open a console window and wait for Creo to be ready. Once Creo is started the application will add a command to the list of commands available in Creo. From now on you can add the command to every toolbar or menu you want.<br/>
For the application it's obligatory to have access to the path ```C:\Temp\```. Here it will store the language files.
