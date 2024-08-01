import cv2
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import pytesseract
from PIL import Image

def process_image(image):
    # white_lo = np.array([128,128,128])
    # white_hi = np.array([150,150,150])
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    lower_gray = np.array([0, 30, 0])
    upper_gray = np.array([180, 100, 255])
    #  HSV = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    # mask = cv2.inRange(image, white_lo, white_hi)
    mask = cv2.inRange(hsv, lower_gray, upper_gray)
    # image[mask>0] = (1,1,1)
    white = np.ones(image.shape, dtype=np.uint8) * 255
    image = np.where(mask[:,:,None] == 255, white, image)

    # lower_gray = np.array([0, 10, 0])
    # upper_gray = np.array([180, 20, 255])
    # Invert the mask to keep non-gray pixels
    # inverted_mask = cv2.bitwise_not(mask)

    # Apply the inverted mask to the original image
    # result = cv2.bitwise_and(image, image, mask=inverted_mask)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    image = cv2.adaptiveThreshold(image,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,\
            cv2.THRESH_BINARY,11,2)
    # ret, binary = cv2.threshold(gray_image, 150, 255, cv2.THRESH_BINARY)
    # return binary
    return image

image_path = 'image copy.png'
image = cv2.imread(image_path)
processed_image = process_image(image)
im_pil = Image.fromarray(processed_image)
im_pil_raw = Image.fromarray(image)

# print(pytesseract.image_to_string(image, lang="eng"))
raw_string = pytesseract.image_to_string(im_pil_raw, lang="eng", config='--psm 13 -c tessedit_char_whitelist=0123456789').replace("\n", '')
print(raw_string)
solution = int(raw_string[:2]) + int(raw_string[-1])  
print(solution) 
im_pil.show()


# cv2.imshow("processed image", image)
# cv2.waitKey(0)
# cv2.destroyAllWindows()