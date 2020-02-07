from PIL import Image
import xlsxwriter
import os

temp_image = '__temp_image'
temp_size = (400, 400)
image_files = ('.jpg', '.jpeg', '.bmp', '.tif', '.tga')

for f in os.listdir('.'):
    if f.endswith(image_files):
        fn, fext = os.path.splitext(f)
        img = Image.open(f)
        wb = xlsxwriter.Workbook(fn + '.xlsx')
        ws = wb.add_worksheet()
        ws.set_column('A:OJ', 1)
        ws.set_default_row(10)
        ws.center_horizontally()
        ws.center_vertically()
        ws.hide_gridlines()
        ws.set_zoom(15)
        ws.fit_to_pages(1,1)

        if max(img.size) > 400:
            img.thumbnail(temp_size, Image.ANTIALIAS)
            img.save(temp_image+fext)
            img = Image.open(temp_image+fext)

        width, height = img.size
        pixels = list(img.getdata())
        pixels = [pixels[i * width:(i + 1) * width] for i in range(height)]
        row, col = 0, 0
        print(f'Generating file {fn:>15}.xlsx -> ', end=' ')
        for sor in pixels:
            for t in sor:
                color = '#%02x%02x%02x' % t if img.mode != 'L' else '#%02x%02x%02x' % (t, t, t)
                cf = wb.add_format()
                cf.set_bg_color(color)
                ws.write_string(row, col, '', cf)
                col += 1
            row += 1
            col = 0
        if os.path.exists(temp_image+fext):
            os.remove(temp_image+fext)
        wb.close()
        print('done!')