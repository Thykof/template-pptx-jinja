import hashlib


from PIL import Image


def get_hash(filename):
    with open(filename, "rb") as f:
        blob = f.read()
    sha1 = hashlib.sha1(blob)
    return sha1.hexdigest()

# See https://github.com/scanny/python-pptx/issues/116
def replace_img_slide(slide, img, img_path):
    # Replace the picture in the shape object (img) with the image in img_path.

    imgPic = img._pic
    imgRID = imgPic.xpath('./p:blipFill/a:blip/@r:embed')[0]
    imgPart = slide.part.related_part(imgRID)

    with open(img_path, 'rb') as f:
        rImgBlob = f.read()

    # replace
    imgPart._blob = rImgBlob
