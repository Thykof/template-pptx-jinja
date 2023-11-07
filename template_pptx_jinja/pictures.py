import hashlib

import pptx

from PIL import Image


def get_hash(filename):
    with open(filename, "rb") as f:
        blob = f.read()
    sha1 = hashlib.sha1(blob)
    return sha1.hexdigest()

# See https://github.com/scanny/python-pptx/issues/116
def replace_img_slide(shape, img_path):
    # Replace the picture in the shape object (img) with the image in img_path.

    new_pptx_img = pptx.parts.image.Image.from_file(img_path)
    slide_part, rId = shape.part, shape._element.blip_rId
    image_part = slide_part.related_part(rId)
    image_part.blob = new_pptx_img._blob
