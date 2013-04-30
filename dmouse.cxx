// DMouse developed by Dhaval Palsana  www.robofreaksindia.wordpress.com

//#include "stdafx.h" //precomplied header files
//#include "windows.h" // Windows-specific
#include "libxl.h"
#include <math.h>
#include <stdio.h>
#include <cv.h>
#include <highgui.h>
#include <cxcore.h>
#include "calibration.cxx"
#include "DMouseConfig.h"
using namespace libxl;
using namespace std;

char welcome(void)
{
    IplImage *welcomeImage = 0;
    welcomeImage = cvLoadImage("images\welcome.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add welcome.jpg to Github
    cvNamedWindow("DMouse- RoboFreaks", CV_WINDOW_AUTOSIZE);
    cvMoveWindow("DMouse- RoboFreaks", 8, 8);
    cvShowImage("DMouse- RoboFreaks", welcomeImage);
    // TODO: Add licence to code (BSD, GPL, ???)
    printf("\nDMouse\nThis software is designed by Dhaval H Palsana\nAll rights reserved-For information and permissions please visit-\nwww.robofreaksindia.wordpress.com");
    while(true)
    {
        int pressedKeyCode = cvWaitKey(1);
        switch(pressedKeyCode & 255)
          {
          case 99: // pressed key == 'c' for 'calibration'
            cvReleaseImage(&welcomeImage);
            cvDestroyAllWindows();
            calibration();
            return 'c';
          case 107: // pressed key == 'k' for 'kill'
            cvReleaseImage(&welcomeImage);
            cvDestroyAllWindows();
            return 'k';
          default:
            continue;
          }
    }
}

int ci, cj;
int bmx[6], gmx[6], rmx[6], bmi[6], gmi[6], rmi[6], hmx[6], smx[6], vmx[6], hmi[6], smi[6], vmi[6];

IplImage* frame = cvCreateImage(cvSize(320, 240), IPL_DEPTH_8U, 3);

int main()
{
    fprintf(stdout,"DMouse Version %d.%d\n",
            DMouse_VERSION_MAJOR,
            DMouse_VERSION_MINOR);

    if (welcome() == 'c')
        return 0;
    int ma[7], mi[7];
    Book* book = xlCreateBook();
    if(book)
    {
        if(book->load(L"dmouse.xls"))
        {
            Sheet* sheet = book->getSheet(0);
            if(sheet)
            {
                int n = 2, d = 1;
            onceagain:
                int maxx = 0, minn = 255;
                for(int i=n; i<=(n + 1); i++)
                {
                    for(int j = 2; j <= 6; j++)
                    {
                        int t;
                        t = sheet->readNum(i, j);
                        maxx = max(maxx, t);
                        minn = min(minn, t);
                    }
                }
                n = n + 2;
                if(n == 16)
                {
                    goto next;
                }
                else
                {
                    ma[d] = maxx;
                    mi[d] = minn;
                    d = d + 1;
                    goto onceagain;
                }
            }
        }
        else
        {
            printf("\nError in reading Database please calibrate once again");
            cvWaitKey(2000);
        }
next:
        book->release();
    }
    SYSTEMTIME st;
    GetSystemTime(&st);
    int year = st.wYear;
    int month = st.wMonth;
    if(year > 2014)
        goto beta;
    float cx = GetSystemMetrics(SM_CXSCREEN);
    float cy = GetSystemMetrics(SM_CYSCREEN);
    printf("\n\ncx = %f\n\ncy = %f\n", cx, cy);
    printf("\n%d %d %d %d %d %d \n%d %d %d %d %d %d ", mi[1], mi[2], mi[3], mi[4], mi[5], mi[6], ma[1], ma[2], ma[3], ma[4], ma[5], ma[6]);
    CvCapture* capture = cvCaptureFromCAM(CV_CAP_ANY);
    if ( !capture )
    {
        fprintf( stderr, "ERROR: capture is NULL \n" );
        getchar();
        return -1;
    }
    IplImage *end = 0;
    end = cvLoadImage("images\end.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: add end.jpg to Github
    cvNamedWindow("DMouse- RoboFreaks", CV_WINDOW_AUTOSIZE);
    cvMoveWindow("DMouse- RoboFreaks", 8, 8);
    cvShowImage("DMouse- RoboFreaks", end);
    IplImage* rgbthresh = cvCreateImage(cvSize(320, 240), 8, 1);
    IplImage* hsvthresh = cvCreateImage(cvSize(320, 240), 8, 1);
    IplImage* thresh = cvCreateImage(cvSize(320, 240), 8, 1);
    int lastX = 0;
    int lastY = 0;
    cvDestroyAllWindows();
    while(1)
    {
        IplImage* img = cvQueryFrame( capture );
        if ( !img )
        {
            fprintf( stderr, "ERROR: Cannot capture form camera. Please ensure that camera is properly connected \n" );
            getchar();
            break;
        }
        cvResize(img, frame, CV_INTER_LINEAR);
        cvFlip(frame, NULL, 1);
        cvInRangeS(frame, cvScalar(mi[1], mi[2] , mi[3]), cvScalar(ma[1], ma[2], ma[3]), rgbthresh);
        cvCvtColor(frame, frame, CV_BGR2HSV);
        cvInRangeS(frame, cvScalar(mi[4], mi[5] , mi[6]), cvScalar(ma[4], ma[5], ma[6]), hsvthresh);
        cvAnd(rgbthresh, hsvthresh, thresh, 0);

        cvSmooth(thresh, thresh, CV_MEDIAN, 7, 0, 0, 0);
        cvShowImage( "mywindow", thresh);
        cvRectangle(thresh, cvPoint(20, 20), cvPoint(300, 220), cvScalar(99, 99, 99), 1, 8, 0);

        CvMoments *moments = (CvMoments*)malloc(sizeof(CvMoments));
        cvMoments(thresh, moments, 1);

        double moment10 = cvGetSpatialMoment(moments, 1, 0);
        double moment01 = cvGetSpatialMoment(moments, 0, 1);
        double area = cvGetCentralMoment(moments, 0, 0);
        static float posX = 0;
        static float posY = 0;
        posX = moment10 / area;
        posY = moment01 / area;
        cvZero(thresh);
        cvCircle(thresh, cvPoint(posX, posY), 9, cvScalar(255), 1, 8, 0);

        float xx = (cx / 14) * ( (posX / 20) - 1);
        float yy = (cy / 10) * ( (posY / 20) - 1);
        int x = xx + 0.5;
        int y = yy + 0.5;
        printf("position (%f,%f)\n", x, y);
        if(x >= 0 && x <= cx && y >= 0 && y <= cy)
        {

            if(abs(x - lastX) >= 1.1)
            {
                if(abs(y - lastY) >= 1.1)
                {
                    SetCursorPos(x, y);
                    printf("position (%d,%d)\n", x, y);
                    lastX = xx;
                    lastY = yy;
                }
            }
        }
        if ( (cvWaitKey(1) & 255) == 27 )
        {
            break;
        }
        cvZero(thresh);
        cvZero(rgbthresh);
        cvZero(hsvthresh);
    }
    cvReleaseCapture(&capture);
    cvReleaseImage(&end);
    cvDestroyAllWindows();
end:
    return 0;

beta:
    IplImage *beta = 0;
    beta=cvLoadImage("images\beta.jpg", CV_LOAD_IMAGE_UNCHANGED);
    cvShowImage("DMouse- RoboFreaks", beta);
    cvWaitKey(0);
    return 0;
}
