import cv2
import numpy as np
import xlwt
import xlrd
import os
from xlutils.copy import copy
from datetime import datetime
from threading import Timer
import time as tt
from datetime import date, time, datetime, timedelta


def point_detect(frame):
    criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 30, 0.001)
    w = 11
    h = 8
    # 储存棋盘格角点的世界坐标和图像坐标对
    objpoints = []  # 在世界坐标系中的三维点
    imgpoints = []  # 在图像平面的二维点
    objp = np.zeros((w * h, 3), np.float32) # 世界坐标系中去掉z坐标变成二维点
    objp[:, :2] = np.mgrid[0:w, 0:h].T.reshape(-1, 2)
    # 储存棋盘格角点的世界坐标和图像坐标对
    img = frame
    ret, corners = cv2.findChessboardCorners(img, (w,h),None)
    gray = cv2.cvtColor(frame,cv2.COLOR_RGB2GRAY)
    if ret == True:
        cv2.cornerSubPix(gray, corners, (11, 11), (-1, -1), criteria)
        objpoints.append(objp)
        imgpoints.append(corners)
        # 将角点在图像上显示
        cv2.drawChessboardCorners(img, (w, h), corners, ret)
        cv2.imshow('img', img)
        cv2.waitKey(3)
    return imgpoints


def write_excel_xls(path, sheet_name, value):
    title = np.array(['row_old','col_old'])  # 生成一个表头
    value = np.vstack((title,value))  # 将表头和数据纵向合并
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, str(value[i][j]))  # 向表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("point_data表格写入数据成功！")


def write_excel_xls_append(path, value):
    title = np.array(['row_new' , 'col_new' ])  # 生成一个表头
    value = np.vstack((title, value))  # 将表头和数据纵向合并
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    cols_old = worksheet.ncols
    new_workbook = copy(workbook) # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i, j+cols_old, str(value[i][j]))  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("point_data表格【追加】写入数据成功！")


def task():
    global t
    # now = datetime.now()
    # period = timedelta(days=0, hours=0, minutes=0, seconds=3)
    # cap_time = now + period
    cap = cv2.VideoCapture(0)
    # i = 0
    while (1):
        ret, frame = cap.read()
        # cv2.imshow("capture", frame)
        cv2.imshow("capture", frame)
        k = cv2.waitKey(1)
        if k == 27:
            cap.release()
            cv2.destroyAllWindows()
            return
        if tt.time() - t > 3:
            imgpoints = point_detect(frame)
            imgpoints = np.array(imgpoints).squeeze()  # 将计算的特征点的数据变形成[88,2]的格式
            book_name_xls = 'point_data.xls'  # 工作簿的名称
            sheet_name_xls = 'point_data表'  # 表的名称
            path = os.path.join(os.getcwd() + '/' + 'point_data.xls')
            p = os.path.exists(path)
            if p == 0:
                write_excel_xls(book_name_xls, sheet_name_xls, imgpoints)
            else:
                write_excel_xls_append(book_name_xls, imgpoints)
            # cv2.imwrite('E:/tempreture/' + str(i) + '.jpg', frame)
            # i += 1
            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            t = tt.time()
        # t = Timer(5, task)  # 间隔inc，调用task函数，参数列表（）
        # t.start()


    # print(imgpoints)
    # print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    # t = Timer(inc, task, (inc, s)).start()   # 间隔inc，调用task函数，参数列表（）
    # t.start()


if __name__ == '__main__':
    t = tt.time()
    task()
    # time.sleep(15)
    # t.cancel()
    # while True:
    #     print(time.time())
    #     time.sleep(5)


