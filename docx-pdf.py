import os
import comtypes.client
import time

format_code = 17

time_start = time.time()

# create the MS word app
word_app = comtypes.client.CreateObject('Word.Application')
word_app.Visible = False
# conversion
file_input = os.path.abspath('C:/Users/Admin/OneDrive/Desktop/VS Projects/sample_file1.doc')
file_output = os.path.abspath('C:/Users/Admin/OneDrive/Desktop/VS Projects/sample_file1.pdf')
word_file = word_app.Documents.Open(file_input)
word_file.SaveAs(file_output,FileFormat=format_code)
word_file.Close()

# close file and application
word_app.Quit()

time_end = time.time()
