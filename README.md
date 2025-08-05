# Magventure TMS Control from Python

  import pytms

# T = pytms.TMS()
Fully control the Magventure TMS at Terminal or from code.
Tested under python 3.8 and 3.13 (need pySerial module).

Hardware requirement: serial port connection between COM2 at TMS machine and the host computer. 
3-wire connection is sufficient: GND->GND, TXD->RXD, RXD->TXD

A USB-to-serial adaptor, like [this one](https://www.amazon.com/Female-Adapter-Chipset-Supports-Windows/dp/B01GA0IZBO/ref=sr_1_10?crid=26ZZRC6MF13A7&dib=eyJ2IjoiMSJ9.ulSsUHaTsJmZ9Jl19PTTci3hFxRjOXORgVD0V2eOceNGoMC92sQkQWfWxMSpTYXjmrIckkqfuhHmZV4ZzdtkTOXU1tbbcNg4rVSvjGA5CQJQB7fskcaLT2lqYDZyUmpBPkkSb7ZdmPrw4H2fL0FM-4ctcz1AFQU6FQ9FITpLqCW8pLZTdoywDmPBfmwW6YiM-LYPK7upLpOLNe-WZrxGzr6gxAtauZc2irazJ5yxCXNKGZK1EzO1V4O12AoPa2MvS8VUZyBbmuieN3_izfBMg0sZceyckAzM5YLUDaqDvVQ.-C1BcM26Jw2HUAaDMnekk0-izmEL1-d5jhVnOIl6tp0&dib_tag=se&keywords=usb+to+usb+crossover+serial+adapter&qid=1744645816&refinements=p_n_feature_six_browse-bin%3A78742982011&rnid=23941269011&s=electronics&sprefix=usb+to+usb+crossover+serial+adapter%2Caps%2C215&sr=1-10), is needed if the computer has no built-in serial port.

# pytms.TMS_GUI()
GUI to control Magventure machine using TMS().

# pytms.rMT()
Estimate resting motor threshold using TMS().
This function also requires to install matplotlib module.

The hardware to record EMG is RTBox. RTBox information can be found [here](https://github.com/xiangruili/RTBox/blob/master/doc/RTBox_v56_user_manual.pdf). 
