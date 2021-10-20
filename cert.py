from PIL import Image, ImageDraw, ImageFont
import pandas as pd
form = pd.read_excel('feedback_form.xlsx')
name_list = form['Name'].to_list()
for i in name_list:
    im = Image.open("demo.jpeg")
    d = ImageDraw.Draw(im)
    location = (215, 882)
    text_color = (0, 137, 209)
    font = ImageFont.truetype("Candara.ttf", 130)
    d.text(location, i, fill=text_color,font=font)
    im.save("certificate_"+i+".pdf")