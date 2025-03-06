from RPLCD.i2c import CharLCD
import time

lcd=CharLCD(i2c_expander='PCF8574', address=0x27, port=1, cols=16, rows=2, dotsize=8)

lcd.clear()
lcd.write_string("Punto de Control")
lcd.cursor_pos = (1, 0)  # Mover a la segunda línea, columna 0
lcd.write_string("AT4580")


    
