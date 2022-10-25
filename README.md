# PowerSaver
This is a power saving project designed to discharge a laptop once the battery is full and begin the charge when it falls below 15%.<br>
Due to changing jobs and varied permissions I've had to write the code in C#, VBA and ASP.net. The project requires the purchase of a USB relay along with any additional cosmetic or safety hardware such as cable, a socket box and USB adapter.

- ## <a href="hardware.md">Hardware requirements and wiring</a>

## Software requirements
The driver for the USB relay is required along with some software to write the code that may include anything from Office, which could utilise Access; Excel; Outlook, or some other IDE like Visual Studio where another language of personal choice could be used.

## <a href="SR_Stage1.md">Stage 1 - VBA on your work/company PC</a>

## <a href="SR_Stage2.md">Stage 2 - Set Up SQL Table</a>

## <a href="SR_Stage3.md">Stage 3 - ASP (Classic to keep things simple)</a>

## <a href="SR_Stage4.md">Stage 4 - Set up the relay</a>

## <a href="SR_Stage5.md">Stage 5 - Set up the VBA code to monitor the database and toggle the relay. In this example I am using Excel</a>

You can adjust this code according to your circumstances. So you may be able to omit a massive chunck of the code if your laptop allows you to install drivers. You could just put the battery monitor on a timer and them call the Init_Com sub when the criteria is met and remove the whole need of a database or a web page etc.<br>

Sometimes the relay sticks. It's only happened once for me. In this case just give it a flick or a tap and it should respond.
