void calibration(void)
{
    int ok = 0;
    printf("\n\nCalibration Started...");
    IplImage *img = 0;
    IplImage *image = 0;
    image = cvLoadImage("start.jpg", CV_LOAD_IMAGE_UNCHANGED);
    cvNamedWindow("Calibrating DMouse- RoboFreaks", CV_WINDOW_AUTOSIZE);
    cvMoveWindow("Calibrating DMouse- RoboFreaks", 400, 80);
    cvShowImage("Calibrating DMouse- RoboFreaks", image);
    cvWaitKey(0);
    image = cvLoadImage("1.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add 1.jpg to Github
    int count = 1;
    cvShowImage("Calibrating DMouse- RoboFreaks", image);
    CvCapture* capture = cvCaptureFromCAM(CV_CAP_ANY);
    if (!capture)
        {
         fprintf(stderr, "ERROR: Cannot capture from camera. Please ensure that camera is properly connected \n");
         getchar();
        }
    cvNamedWindow("From Camera", CV_WINDOW_AUTOSIZE);
    cvMoveWindow("From Camera", 50, 80);
    HWND hChild;
    hChild = (HWND)cvGetWindowHandle("Calibrating DMouse- RoboFreaks");
    while (true)
    {
        IplImage* img = cvQueryFrame(capture);
        if (!img)
        {
            fprintf(stderr, "ERROR: frame is null...\n");
            getchar();
            break;
        }
        cvResize(img, frame, CV_INTER_LINEAR);
        int height = frame->height;
        int width = frame->width;
        cvFlip(frame, NULL, 1);
        CvScalar s;
        s.val[2] = 0;
        s.val[1] = 0;
        s.val[0] = 255;
        cvCircle(frame, cvPoint(20, 20), 9, s, 2, 9);
        cvCircle(frame, cvPoint(300, 20), 9, s, 2, 9);
        cvCircle(frame, cvPoint(20, 220), 9, s, 2, 9);
        cvCircle(frame, cvPoint(300, 220), 9, s, 2, 9);
        cvCircle(frame, cvPoint(160, 120), 9, s, 2, 9);
        cvShowImage("From Camera", frame);
        SetForegroundWindow(hChild);

        if (count == 1)////////////////////////////////////////////////////////FIRST
        {
            ci = 20;
            cj = 20;
            calibrun(count);
            if(ok == 1)
            {
                ok = 0;
                count = 2;
                image = cvLoadImage("2.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add 2.jpg to Github
                cvShowImage("Calibrating DMouse- RoboFreaks", image);
            }
        }

        if (count == 2)//////////////////////////////////////////////////////////SECOND
        {
            ci = 20;
            cj = 300;
            calibrun(count);
            if(ok == 1)
            {
                ok = 0;
                count = 3;
                image = cvLoadImage("3.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add 3.jpg to Github
                cvShowImage("Calibrating DMouse- RoboFreaks", image);
            }
        }

        if (count == 3)///////////////////////////////////////////////////////////THIRD
        {
            ci = 220;
            cj = 20;
            calibrun(count);
            if(ok == 1)
            {
                ok = 0;
                count = 4;
                image = cvLoadImage("4.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add 4.jpg to Github
                cvShowImage("Calibrating DMouse- RoboFreaks", image);
            }
        }

        if (count == 4)///////////////////////////////////////////////////////////FOURTH
        {
            ci = 220;
            cj = 300;
            calibrun(count);
            if(ok == 1)
            {
                ok = 0;
                count = 5;
                image = cvLoadImage("5.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add 5.jpg to Github
                cvShowImage("Calibrating DMouse- RoboFreaks", image);
            }
        }

        if (count == 5)///////////////////////////////////////////////////////////FIFTH
        {
            ci = 120;
            cj = 160;
            calibrun(count);
            if(ok == 1)
            {
                ok = 0;
                count = 6;
                image = cvLoadImage("done.jpg", CV_LOAD_IMAGE_UNCHANGED); // TODO: Add done.jpg to Github
                cvShowImage("Calibrating DMouse- RoboFreaks", image);
            }
        }
        if ((cvWaitKey(1) & 255) == 27)
        {
            break;
        }
    }
        Book* book = xlCreateBook();
        if(book)
        {
            Sheet* sheet = book->addSheet(L"Sheet1");
            if(sheet)
            {
                sheet->writeStr(1, 2, L"Point 1");
                sheet->writeStr(1, 3, L"Point 2");
                sheet->writeStr(1, 4, L"Point 3");
                sheet->writeStr(1, 5, L"Point 4");
                sheet->writeStr(1, 6, L"Point 5");

                sheet->writeStr(2, 1, L"Blue");
                sheet->writeNum(2, 2, bmx[1]);
                sheet->writeNum(2, 3, bmx[2]);
                sheet->writeNum(2, 4, bmx[3]);
                sheet->writeNum(2, 5, bmx[4]);
                sheet->writeNum(2, 6, bmx[5]);

                sheet->writeNum(3, 2, bmi[1]);
                sheet->writeNum(3, 3, bmi[2]);
                sheet->writeNum(3, 4, bmi[3]);
                sheet->writeNum(3, 5, bmi[4]);
                sheet->writeNum(3, 6, bmi[5]);

                sheet->writeStr(4, 1, L"Green");
                sheet->writeNum(4, 2, gmx[1]);
                sheet->writeNum(4, 3, gmx[2]);
                sheet->writeNum(4, 4, gmx[3]);
                sheet->writeNum(4, 5, gmx[4]);
                sheet->writeNum(4, 6, gmx[5]);

                sheet->writeNum(5, 2, gmi[1]);
                sheet->writeNum(5, 3, gmi[2]);
                sheet->writeNum(5, 4, gmi[3]);
                sheet->writeNum(5, 5, gmi[4]);
                sheet->writeNum(5, 6, gmi[5]);

                sheet->writeStr(6, 1, L"Red");
                sheet->writeNum(6, 2, rmx[1]);
                sheet->writeNum(6, 3, rmx[2]);
                sheet->writeNum(6, 4, rmx[3]);
                sheet->writeNum(6, 5, rmx[4]);
                sheet->writeNum(6, 6, rmx[5]);

                sheet->writeNum(7, 2, rmi[1]);
                sheet->writeNum(7, 3, rmi[2]);
                sheet->writeNum(7, 4, rmi[3]);
                sheet->writeNum(7, 5, rmi[4]);
                sheet->writeNum(7, 6, rmi[5]);

                sheet->writeStr(8, 1, L"Hue");
                sheet->writeNum(8, 2, hmx[1]);
                sheet->writeNum(8, 3, hmx[2]);
                sheet->writeNum(8, 4, hmx[3]);
                sheet->writeNum(8, 5, hmx[4]);
                sheet->writeNum(8, 6, hmx[5]);

                sheet->writeNum(9, 2, hmi[1]);
                sheet->writeNum(9, 3, hmi[2]);
                sheet->writeNum(9, 4, hmi[3]);
                sheet->writeNum(9, 5, hmi[4]);
                sheet->writeNum(9, 6, hmi[5]);

                sheet->writeStr(10, 1, L"Sat");
                sheet->writeNum(10, 2, smx[1]);
                sheet->writeNum(10, 3, smx[2]);
                sheet->writeNum(10, 4, smx[3]);
                sheet->writeNum(10, 5, smx[4]);
                sheet->writeNum(10, 6, smx[5]);

                sheet->writeNum(11, 2, smi[1]);
                sheet->writeNum(11, 3, smi[2]);
                sheet->writeNum(11, 4, smi[3]);
                sheet->writeNum(11, 5, smi[4]);
                sheet->writeNum(11, 6, smi[5]);

                sheet->writeStr(12, 1, L"Val");
                sheet->writeNum(12, 2, vmx[1]);
                sheet->writeNum(12, 3, vmx[2]);
                sheet->writeNum(12, 4, vmx[3]);
                sheet->writeNum(12, 5, vmx[4]);
                sheet->writeNum(12, 6, vmx[5]);

                sheet->writeNum(13, 2, vmi[1]);
                sheet->writeNum(13, 3, vmi[2]);
                sheet->writeNum(13, 4, vmi[3]);
                sheet->writeNum(13, 5, vmi[4]);
                sheet->writeNum(13, 6, vmi[5]);
            }
            book->save(L"dmouse.xls");
            book->release();
        }
    cvDestroyAllWindows();
    cvReleaseCapture( &capture );
    cvDestroyWindow( "From Camera" );
    cvReleaseImage(&frame);
}

void calibrun(int count)
{
    CvScalar s;
    int i, j;
    int mxb = 0, mxg = 0, mxr = 0, mxh = 0, mxs = 0, mxv = 0;
    int mib = 255, mig = 255, mir = 255, mih = 255, mis = 255, miv = 255;
    int r, g, b, h, sa, v;
    if((cvWaitKey(1) & 255) == 107) // key pressed == 'k'
    {
        for(i = (ci - 2); i <= (ci + 2); i++)
            {
                for(j = (cj - 2); j <= (cj + 2); j++)
                {
                    s = cvGet2D(frame, i, j);
                    b = s.val[0];
                    g = s.val[1];
                    r = s.val[2];
                    mxb = max(mxb, b);
                    mxg = max(mxg, g);
                    mxr = max(mxr, r);
                    mib = min(mib, b);
                    mig = min(mig, g);
                    mir = min(mir, r);
                }
            }
            cvCvtColor(frame, frame, CV_BGR2HSV);
            for(i = (ci - 2); i <= (ci+2); i++)
            {
                for(j = (cj - 2); j <= (cj + 2); j++)
                {
                    s = cvGet2D(frame, i, j);
                    h = s.val[0];
                    sa = s.val[1];
                    v = s.val[2];
                    mxh = max(mxh, h);
                    mxs = max(mxs, sa);
                    mxv = max(mxv, v);
                    mih = min(mih, h);
                    mis = min(mis, sa);
                    miv = min(miv, v);
                }
            }
            printf("\nb: %d %d", mxb, mib);
            printf("\ng: %d %d", mxg, mig);
            printf("\nr: %d %d", mxr, mir);
            printf("\nh: %d %d", mxh, mih);
            printf("\ns: %d %d", mxs, mis);
            printf("\nv: %d %d", mxv, miv);
            printf("\n%d Down.....", count);
            ok = 1;
    }
    bmx[count] = mxb;
    gmx[count] = mxg;
    rmx[count] = mxr;
    bmi[count] = mib;
    gmi[count] = mig;
    rmi[count] = mir;
    hmx[count] = mxh;
    smx[count] = mxs;
    vmx[count] = mxv;
    hmi[count] = mih;
    smi[count] = mis;
    vmi[count] = miv;
}
