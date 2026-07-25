import pyvisa

rm = pyvisa.ResourceManager()
visa_addr = 'USB0::0x2A8D::0x2404::MY65150204::INSTR'

try:
    nst = rm.open_resource(visa_addr)
    nst.timeout = 5000  
    
    print("Connected to --> " + nst.query("*IDN?").strip())
    
    # 1. Reset instrument and clear the error queue
    nst.write("*RST")
    nst.write("*CLS")
    
    print("\nSending ARB SIN commands to the instrument...")
    
    # 2. Send the exact ARB commands you found
    nst.write(":SOUR:FUNC:MODE ARB")
    nst.write(":SOUR:ARB:FUNC:SHAP SIN")
    nst.write(":SOUR:ARB:CURR:LEV 0.040")  # Test with 40mA amplitude
    nst.write(":SOUR:ARB:FREQ 1")          # Test with 1 Hz frequency
    
    # 3. Ask the instrument if it accepted the commands or threw an error
    error_response = nst.query(":SYST:ERR?")
    print(f"\nInstrument Error Queue returns: {error_response.strip()}")
    
    # 4. Interpret the result
    if "+0," in error_response:
        print("✅ SUCCESS: Your instrument accepted the ARB commands! The feature is working.")
    else:
        print("❌ FAILED: The instrument threw an error. It does not recognize or support these ARB commands.")

except pyvisa.VisaIOError as e:
    print(f"VISA Error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    try:
        nst.close()
        rm.close()
        print("\nConnection safely closed.")
    except:
        pass