import os
import xlsxwriter

def copy_source_image(research_data_folder,image_folder_name,image_name):
    os.system('cd '+research_data_folder+'\\'+image_folder_name+
              ' &&copy '+image_name+'.tif out\\'+image_folder_name+'-0000.tif')

def text_cell_format(wb):
    cell_format = wb.add_format()
    cell_format.set_text_wrap()
    return cell_format

def createXLSX(research_data_folder,image_folder_name,image_name):
    workbook = xlsxwriter.Workbook(research_data_folder+'\\'+image_folder_name+
                                   '\\image_sharp_parameter.xlsx')
    worksheet = workbook.add_worksheet('Sheet1')


    worksheet.row_sizes
    worksheet.write('A1', 'imgName')
    worksheet.write('B1', "Radius")
    worksheet.write('C1', "Amount")
    copy_source_image(research_data_folder,image_folder_name,image_name)
    worksheet.write('A2',image_folder_name+'-0000')
    worksheet.write_number('B2',0)
    worksheet.write_number('C2', 0)
    return workbook






def main():
    image_name = input('Please enter image name:')
    image_folder_name = image_name.replace('-', '')
    book_name = image_name[0:image_name.index('-')]
    research_data_folder = 'C:\\Users\\Administrator\\DictaProg Dropbox\\OCR Library\\Dicta Library' \
                           '\\research-data'
    cd_research_data_folder_command='cd '+research_data_folder
    create_image_folder_command='md '+image_folder_name
    copy_source_image_command='copy '+book_name + '\\out\\' + image_name +'.tif ' + image_folder_name \
                              + '\\'+ image_name + '.tif'
    create_image_out_folder_command='md '+ image_folder_name + '\\out'
    cd_image_folder_command=cd_research_data_folder_command+'\\'+image_folder_name
    os.system(cd_research_data_folder_command + '&&'+create_image_folder_command+'&&'+
              copy_source_image_command+'&&md ' + create_image_out_folder_command)

    page_sharp_workbook = createXLSX(research_data_folder, image_folder_name, image_name)
    page_sharp_sheet = page_sharp_workbook.get_worksheet_by_name('Sheet1')
    for j in range(1, 6):
        for i in range(1, 6):
            print('Radius: ' + str(j * 6) + '\nAmount: ' + str(i * 3))
            page_sharp_sheet.write(((j - 1) * 5) + i, 0,
                                   image_folder_name + '-' + ("{:02d}".format(j * 6)) + ("{:02d}".format(i * 3)))
            page_sharp_sheet.write_number(((j - 1) * 5) + i, 1, j * 6)
            page_sharp_sheet.write_number(((j - 1) * 5) + i, 2, i * 3)
            os.system(cd_image_folder_command+ '&&magick convert ' + image_name + '.tif'
                + ' -sharpen ' + str(j * 6) + 'x' + str(i * 3) + ' out\\' + image_folder_name +
                '-' + ("{:02d}".format(j * 6)) + ("{:02d}".format(i * 3)) + '.tif')
    os.system('plink ubuntu@54.76.178.38 -i C:/Users/Administrator/Desktop/DictaTestingKeyPair.ppk'
              ' cd ~/\'DictaProg Dropbox\'/\'OCR Library\'/\'Dicta Library\'/research-data;'
              'python3 ~/ocr-research/scripts/main.py ' + image_folder_name + ' --all')
    page_sharp_workbook.close()



if __name__ == "__main__":
    main()

