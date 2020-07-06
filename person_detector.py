"""problem: Design a system for person detection and logging

working:
1. read the video stream of webcam or any video from internet using python
2. perform person detection in a video stream
3. if person is detected, crop the person box only(not entire image)
4. store the cropped image in the xlsx file along with the the timestamp
5. only single entry should be there for each person. you can use tracker or use your own logic to implement the task.
    worksheet.write(row, 0, filename)

note:
dont just copy paste the code. it wont be accepted for evaluation
comments explaining the working of each definition and library specific lines are mandatory
try to change the fixed values in the referred code and document the changes in output if any"""
# date :July 1 2020
import glob
import os
import shutil
from datetime import datetime
import cv2
import imutils
import xlsxwriter


def main():
    # Create a VideoCapture object and read from input file
    cap = cv2.VideoCapture(video_path)

    # Haar is better for human face. Hog with SVM is classic for human detection
    hog = cv2.HOGDescriptor()

    # Set the support vector machine to be pre-trained for people detection
    hog.setSVMDetector(cv2.HOGDescriptor_getDefaultPeopleDetector())

    # workbook = xlsxwriter.Workbook(excel_file_name + ".xlsx")
    # worksheet = workbook.add_worksheet()
    #
    # # Widen the first column to make the text clearer.
    # worksheet.set_column('A:A', 25)
    # worksheet.set_column('B:B', 25)
    #
    # worksheet.write(0, 0, 'timestamp')
    # worksheet.write(0, 1, 'detected peoples')


    # Check if camera opened successfully
    if not cap.isOpened():
        print("Error opening video  file")

    # Read until video is completed
    while cap.isOpened():



        # Capture frame-by-frame
        ret, frame = cap.read()
        if ret:
            frame = imutils.resize(frame, width=min(500, frame.shape[0]))
            # Display the resulting frame
            #cv2.imshow('Frame', frame)

            """ 
            Detect people in the image
            1)winStride: It is the step size in the x and y direction of our sliding window.
            2)padding : It is a tuple which indicates the number of pixels in both the x and y direction in which the sliding window ROI is “padded” prior to HOG feature extraction.
            3)scale : To control the scale of the image pyramid (allowing us to detect people in images at multiple scales) 
            """

            rect, weight = hog.detectMultiScale(frame, winStride=(5, 9), padding=(32, 32), scale=1.5)

            # call draw_detections function to draw boundaries to  detected person
            draw_detections(frame, rect)

            # Press Q on keyboard to  exit
            if cv2.waitKey(25) & 0xFF == ord('q'):
                break

        # Break the loop
        else:
            break

    # When everything done, release the video capture object
    cap.release()

    # Closes all the frames
    cv2.destroyAllWindows()


def draw_detections(frame, rect,thickness=1):
    """this function draws rectangle on the image where the person is detected and saves the croped image"""

    global row
    for x, y, w, h in rect:
        # row : to store the image on paticular row in xlsx sheet
        row = row + 10

        # the HOG detector returns slightly larger rectangles than the real objects.
        # so we slightly shrink the rectangles to get a nicer output.
        pad_w, pad_h = int(0.15 * w), int(0.05 * h)

        # draw rectangle
        cv2.rectangle(frame, (x + pad_w, y + pad_h), (x + w - pad_w, y + h - pad_h), (0, 255, 0), thickness)

        # show region where rectangle is drawn
        cv2.imshow("detected person", frame[y + pad_h: y + h - pad_h, x + pad_w: x + w - pad_w])

        cv2.imshow("video source", frame)
        # retrive the date and time at the moment function now() is called
        filename = datetime.now().strftime("%d-%m-%Y_%I:%M:%S_%p")

        # save the image the filename of timestamp
        cv2.imwrite(images_dir+filename+".png", frame[y + pad_h: y + h - pad_h, x + pad_w: x + w - pad_w])

def xls(excel_file_name):
    """this function creates workbook and worksheet and the dump the data in sheet"""
    # Enter the path where image is saved

    images = glob.glob(images_path)
    workbook = xlsxwriter.Workbook(excel_file_name+'.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 25)
    worksheet.set_column('B:B', 25)

    # Here height is 30
    worksheet.set_default_row(30)

    worksheet.write(0, 0, 'timestamp')
    worksheet.write(0, 1, 'detected peoples')

    image_row = 1
    image_col = 1
    for image in images:
        date = os.path.basename(image)
        date = os.path.splitext(date)[0]
        # Insert an image with scaling.
        worksheet.write(image_row, image_col-1, date)

        worksheet.insert_image(image_row, image_col, image, {'x_scale': 0.7, 'y_scale': 0.7,
                                                             'x_offset': 15, 'y_offset': 1, 'positioning': 1})

        image_row += 3
    workbook.close()
    # Enter the path of folder which need to be deleted
    shutil.rmtree(images_dir)
    os.makedirs(images_dir)



row = 0
# path = "/home/tushar/Downloads/180301_06_B_CityRoam_01.mp4"
video_path = input("please enter the full path of video source : \n")

# excel_file_name = "test"
excel_file_name = input("\nBy which name you want to generate excel file...(please do not add extension): \n")

print("\n please wait till the window closes...")

#directory path of the images
images_dir = "/home/tushar/Desktop/training/support/Tushar/person detection/images/"

#path of directory where images will to be saved
images_path = '/home/tushar/Desktop/training/support/Tushar/person detection/images/*.png'

main()

xls(excel_file_name)