from pptx import Presentation
from pptx.util import Inches
import os
def addBulletSlide(ppt_name = None):
    if ppt_name:
        ppt_name = ppt_name + '.pptx'
    prs = Presentation(ppt_name)
    heading = input('Please Enter heading of Bullet Slide:-- ')
    first_text = input('Please Enter line in heading:-- ')
    bullet_line = input('Please Enter bullet line:-- ')
    bullet_slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]
    title_shape.text = heading
    tf = body_shape.text_frame
    tf.text = first_text
    p = tf.add_paragraph()
    p.text = bullet_line
    p.level = 1
    p = tf.add_paragraph()
    p.text = ''
    p.level = 2
    return prs

def addImageSlide(img_path, ppt_name = None):
    if ppt_name:
        ppt_name = ppt_name + '.pptx'
    prs = Presentation(ppt_name)
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    left =Inches(1.6)
    top = Inches(1)
    pic = slide.shapes.add_picture(img_path, left, top)
    return prs

def createppt(ppt_name, *args):
    #1 create text slide
    x = None
    decision = input('Want to add bullet slide:-- ')
    if decision == 'yes' or decision == 'Yes':
        x = 1
    else:
        x = None
    if x:
        try:
            prs = Presentation(ppt_name + '.pptx')
            prs = addBulletSlide(ppt_name)
            prs.save(ppt_name + '.pptx')
        except:
            prs = addBulletSlide()
            prs.save(ppt_name + '.pptx')

    #2 create image slide
    for i in args:
        img_path = str(i) + '.png'
        try:
            prs = Presentation(ppt_name + '.pptx')
            prs = addImageSlide(img_path, ppt_name)
            prs.save(ppt_name + '.pptx')
            os.remove(img_path)
        except:
            prs = addImageSlide(img_path)
            prs.save(ppt_name + '.pptx')
            os.remove(img_path)
            

def makeppt(file_name, ppt_name):
    import matplotlib.pyplot as plt
    #%matplotlib inline
    try:
        with open(str(file_name)) as f:
            data = f.read()
            l = []
            for i in range(2):
                n1 = data.find('[')
                n2 = data.find(']')
                d = data[n1 + 1: n2]
                l.append([int(i) for i in d if i.isdigit()])
                data = data[n2+1:]
    except:
        return None
    x, y = l[0], l[1]
    plt.title('X vs Y dashed line plot ')
    plt.xlabel('x ---------------->')
    plt.ylabel('y ---------------->')
    plt.plot(x, y, color='green', marker='o', linestyle='dashed',
        linewidth=2, markersize=12)
    plt.grid()
    plt.savefig('plot.png', dpi = 100, pad_inches = 0.3)
    plt.close()
    plt.title('X vs Y Scatter plot')
    plt.xlabel('x ---------------->')
    plt.ylabel('y ---------------->')
    g = plt.scatter(x, y,color = ['r', 'g'], marker = 'o', edgecolor = 'orange', s = 150)
    plt.grid()
    plt.savefig('scatter.png', dpi = 100, pad_inches = 0.3)
    plt.close()
    return createppt(ppt_name, 'plot', 'scatter')
file_name = input('Enter the name of text file:-- ')
if '.txt' in file_name:
    file_name = file_name
else:
    file_name = file_name + '.txt'
name = input('enter name of pptx file:-- ')
makeppt(file_name, name)
