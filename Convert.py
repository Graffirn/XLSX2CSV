import pandas as pd
import os
import glob

for xlsxFile in glob.glob(os.path.join('.\\', '*.xlsx')):
  # Create contain folder
  try:
    os.mkdir(".\\" + xlsxFile[2:-5])
    print("Directory " , xlsxFile[2:-5] ,  " created ") 
  except FileExistsError:
    print("Directory " , xlsxFile[2:-5] ,  " already exists")
    file = glob.glob(xlsxFile[2:-5] + "\\*")
    for f in file:
      os.remove(f)
    print('Data in ' + xlsxFile[2:-5] + ' has been removed.')
  # Start converting
  excel = pd.read_excel(xlsxFile[2:], None)
  for name, data in excel.items():
    data.to_csv(xlsxFile[2:-5] + '\\' + name + '.csv', encoding = 'utf-8-sig', index = False)
  print("Done!")