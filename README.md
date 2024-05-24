To run the auto email generator:
1. Must have python installed
2. Open a command window and enter the following lines to ensure you have correct dependencies:

   python --version
   
   pip install pyperclip
   
   python -m pip install pywin32

3. You will have to change in the working directory in 'gen_gui.py' (should be on line 84) to match the directory that you have the project downloaded to 

4. If all dependencies are met, run the project by double clicking 'gen_gui.py' to run it
5. A small GUI window should pop up asking for a name
6. After entering a name and pressing Enter, the GUI should close and a new Outlook email should open 
