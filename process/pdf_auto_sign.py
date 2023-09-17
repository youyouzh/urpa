import io
import os.path

from PIL import Image, ImageDraw, ImageFont
from fitz import Rect
from pyrect import Box

from handright import Template, handwrite
from itertools import product
from base.log import logger
# pip install pymupdf
import fitz


# 汉仪新蒂绿豆体Senty Pea Handwriting
HANDWRITE_FONT = ImageFont.truetype(r'data\SentyPea.ttf', size=100)
# HANDRIGHT_FONT = ImageFont.truetype(r'data\叶根友钢笔行书.ttf', size=100)


# 去除pdf的水印
def remove_pdf_watermark(pdf_filepath: str, watermark_rgb):
    # 打开源pfd文件
    pdf_file = fitz.open(pdf_filepath)
    save_page_dir = r'temp'
    if not os.path.isdir(save_page_dir):
        os.makedirs(save_page_dir)

    logger.info('read pdf file success: {}, page size: {}'.format(pdf_filepath, len(pdf_file)))
    page_index = 0
    # page在pdf文件中遍历
    for page in pdf_file:
        # 获取每一页对应的图片pix (pix对象类似于我们上面看到的img对象，可以读取、修改它的 RGB)
        # page.get_pixmap() 这个操作是不可逆的，即能够实现从 PDF 到图片的转换，但修改图片 RGB 后无法应用到 PDF 上，只能输出为图片
        pixmap = page.get_pixmap()

        # 遍历图片中的宽和高，如果像素的rgb值总和大于510，就认为是水印，转换成255，255,255-->即白色
        for pos in product(range(pixmap.width), range(pixmap.height)):
            if sum(pixmap.pixel(pos[0], pos[1])) >= sum(watermark_rgb):
                pixmap.set_pixel(pos[0], pos[1], (255, 255, 255))

        # 保存去掉水印的图片
        pixmap.pil_save(os.path.join(save_page_dir, 'pdf-{}.png'.format(page_index)), dpi=(30000, 30000))
        logger.info('去除水印完成：页码： {}'.format(page_index))
        page_index += 1


# 去除的pdf水印添加到pdf文件中
def merge_image_to_pdf(image_dir: str, generate_pdf_path: str):
    # 水印截图所在的文件夹
    generate_pdf_path = generate_pdf_path if generate_pdf_path else 'merge.pdf'
    pdf = fitz.open(generate_pdf_path)
    # 图片数字文件先转换成int类型进行排序
    image_files = sorted(os.listdir(image_dir), key=lambda x: int(str(x).split('.')[0]))
    for image_file in image_files:
        image_pdf = fitz.open(os.path.join(image_dir, image_file))
        # 将打开后的图片转成单页pdf
        pdf_bytes = image_pdf.convertToPDF()
        pdf_page = fitz.open("pdf", pdf_bytes)
        # 将单页pdf插入到新的pdf文档中
        pdf.insertPDF(pdf_page)
    pdf.save("源码找落落阿_完成.pdf")
    pdf.close()


def generate_sign_text_image(sign_text: str):
    # 创建一个空白的图片
    img = Image.new('RGB', (500, 100), color=(255, 255, 255))
    # 基于该图片绘制一个画布
    draw = ImageDraw.Draw(img)
    # 在图片上绘制汉字
    draw.text((10, 10), sign_text, font=HANDWRITE_FONT, fill=(0, 0, 0))
    return img


def sign_text_on_image(background_image, sign_text: str):
    # 基于该图片绘制一个画布
    draw = ImageDraw.Draw(background_image)
    # 设置字体和大小
    # 在图片上绘制汉字
    draw.text((10, 10), sign_text, font=HANDWRITE_FONT, fill=0)
    background_image.show()


def sign_handright_text_on_image(background_image: Image, sign_text: str, rect=Rect(50, 50, 100, 100), scale=True):
    # 对图片进行放大后在签名，字体太小签名锯齿比较大
    scale_size = 1
    if scale:
        width, height = background_image.size
        scale_size = 2400 / width
        background_image = background_image.resize((int(width * scale_size),
                                                    int(height * scale_size)), resample=Image.LANCZOS)
    template = Template(
        background=background_image,
        font=HANDWRITE_FONT,
        line_spacing=100,
        fill=0,  # 字体“颜色”
        left_margin=rect.top_left.x,
        top_margin=rect.top_left.y,
        right_margin=0,
        bottom_margin=0,
        word_spacing=5,
        line_spacing_sigma=0.5,  # 行间距随机扰动
        font_size_sigma=0.5,  # 字体大小随机扰动
        word_spacing_sigma=5,  # 字间距随机扰动
        word_spacing_decrease=35,
        start_chars="“（[<",  # 特定字符提前换行，防止出现在行尾
        end_chars="，。",  # 防止特定字符因排版算法的自动换行而出现在行首
        perturb_x_sigma=0.5,  # 笔画横向偏移随机扰动
        perturb_y_sigma=2,  # 笔画纵向偏移随机扰动
        perturb_theta_sigma=0.005,  # 笔画旋转偏移随机扰动，大于0.5会导致空白点
    )
    handwrite_images = handwrite(sign_text, template)
    handwrite_image: Image = next(handwrite_images)
    # 还原图片
    if scale:
        width, height = handwrite_image.size
        # handwrite_image.thumbnail((int(width / scale_size), int(height / scale_size)), resample=Image.LANCZOS)
    return handwrite_image


# 测试手写字体效果
def test_handwriting():
    background_image = Image.new(mode="1", size=(2400, 800), color=1)  # 白底黑字
    # background_image = Image.new(mode="RGBA", size=(2000, 800), color=(0, 0, 0, 1))
    sign_text = '信银理财有限责任公司（代表“信银理财安盈象固收稳利十四个月封闭式31号理财产品”）'
    sign_handright_text_on_image(background_image, sign_text, scale=False)
    # sign_text_on_image(background_image, sign_text)
    # sign_text_on_image(sign_text)


def show_pdf_page(pdf_page):
    pixmap = pdf_page.get_pixmap()
    img = Image.open(io.BytesIO(pixmap.tobytes()))
    # 显示图像
    img.show()


# 自动签名
def auto_sign_pdf(source_pdf_filepath, sign_page_index, sign_rect, sign_text):
    if not os.path.isfile(source_pdf_filepath):
        logger.info('The source pdf file is not exist: {}'.format(source_pdf_filepath))
        return False

    source_pdf = fitz.open(source_pdf_filepath)
    logger.info('read pdf file success: {}, page size: {}'.format(source_pdf, len(source_pdf)))
    # 需要签名的页
    sign_pdf_page = source_pdf[sign_page_index]
    # 转成PIL-Image图片之后再进行签名，直接在PDF上签名会有各种旋转和坐标问题
    sign_pdf_image = Image.open(io.BytesIO(sign_pdf_page.get_pixmap().tobytes()))
    signed_image = sign_handright_text_on_image(sign_pdf_image, sign_text, sign_rect)
    signed_image.show()

    sign_image = generate_sign_text_image(sign_text)
    # 转成字节流
    sign_image_byte_arr = io.BytesIO()
    sign_image.save(sign_image_byte_arr, format='PNG')
    sign_image_byte_arr = sign_image_byte_arr.getvalue()

    # if not sign_page.is_wrapped:
    #     sign_page.wrap_contents()

    # sign_page.insert_image(sign_rect, stream=sign_image_byte_arr)
    # show_pdf_page(sign_page)
    # source_pdf.save(r'E:\需求相关资料\需求-PDF自动签名\sign.pdf')


def test_auto_sign_pdf():
    auto_sign_pdf(r'data\匠心利率2号风险揭示书.pdf',
                  9,
                  Rect(0, 500, 600, 600),
                  '信银理财有限责任公司（代表“信银理财安盈象固收稳利十四个月封闭式31号理财产品”）')


if __name__ == '__main__':
    # test_handwriting()
    test_auto_sign_pdf()
