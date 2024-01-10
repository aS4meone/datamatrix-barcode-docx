from pystrich.datamatrix import DataMatrixEncoder
from PIL import Image
import io
import matplotlib.pyplot as plt


def generate_datamatrix(data):
    encoder = DataMatrixEncoder(data)
    barcode_data = encoder.get_imagedata()
    return Image.open(io.BytesIO(barcode_data))


data = "069424124121418605\"7U/G0?"
barcode_image = generate_datamatrix(data)

plt.imshow(barcode_image, cmap='gray')
plt.axis('off')
plt.show()
