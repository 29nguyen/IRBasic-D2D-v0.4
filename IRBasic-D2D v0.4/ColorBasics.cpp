//------------------------------------------------------------------------------
// <copyright file="ColorBasics.cpp" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

//std::string test = std::to_string(iSensorCount);
//MessageBox(NULL, LPCWSTR(test.c_str()), L"Result", MB_OK);

//	OPENFILENAME ofn;
//	char szFileName[MAX_PATH] = "";
//	ZeroMemory(&ofn, sizeof(ofn));
//	ofn.lStructSize = sizeof(OPENFILENAME);
//	ofn.lpstrFilter = (LPCWSTR)L"BMP file (*.bmp)\0*.bmp\0All Files (*.*)\0*.*\0";
//	ofn.lpstrFile = (LPWSTR)szFileName;
//	ofn.nMaxFile = MAX_PATH;
//	ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
//	GetSaveFileName(&ofn);
//	m_str = ofn.lpstrFile;

#include "stdafx.h"
#include <strsafe.h>
#include "ColorBasics.h"
#include "resource.h"
#include <iostream>
#include <string>
#include <commdlg.h>
#include <windows.h>

/// <summary>
/// Entry point for the application
/// </summary>
/// <param name="hInstance">handle to the application instance</param>
/// <param name="hPrevInstance">always 0</param>
/// <param name="lpCmdLine">command line arguments</param>
/// <param name="nCmdShow">whether to display minimized, maximized, or normally</param>
/// <returns>status</returns>

/// Dialog to specify the number of devices
INT_PTR CALLBACK NumberDeviceDlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	static ChooseNumberImage *nParam = NULL;
	switch (uMsg)
	{
	case WM_INITDIALOG:
		nParam = (ChooseNumberImage*)lParam;
		SetDlgItemInt(hwnd, IDC_NUM_IMAGE, 1, FALSE);	//	Set the initial value to 1 (one image)

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
		{
			BOOL bSuccess;
			int nTimes = GetDlgItemInt(hwnd, IDC_NUM_IMAGE, &bSuccess, FALSE);
			if (bSuccess)
			{
				nParam->selection = (UINT32)nTimes;
			}
			EndDialog(hwnd, LOWORD(wParam));
			return TRUE;
		}
		case IDCANCEL:
		{
			EndDialog(hwnd, LOWORD(wParam));
			return TRUE;
		}
		break;
		}
	}
	return FALSE;

}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    CColorBasics application;
    application.Run(hInstance, nCmdShow);
}

/// <summary>
/// Constructor
/// </summary>
CColorBasics::CColorBasics() :
	m_pD2DFactory(NULL),
	m_pDrawColor(NULL),
	m_hNextColorFrameEvent(INVALID_HANDLE_VALUE),
	m_pColorStreamHandle(INVALID_HANDLE_VALUE),
	m_bSaveScreenshot(false),
	m_bOffIrEmitter(false),
	m_continuouscap(false),
	m_pNuiSensor(NULL),
	m_pTempColorBuffer(NULL),

	m_pD2DFactory_1(NULL),
	m_pDrawColor_1(NULL),
	m_hNextColorFrameEvent_1(INVALID_HANDLE_VALUE),
	m_pColorStreamHandle_1(INVALID_HANDLE_VALUE),
	m_bSaveScreenshot_1(false),
	m_pNuiSensor_1(NULL),
	m_pTempColorBuffer_1(NULL)
{
}

/// <summary>
/// Destructor
/// </summary>
CColorBasics::~CColorBasics()
{
    if (m_pNuiSensor)
    {
        m_pNuiSensor->NuiShutdown();
    }

	if (m_pNuiSensor_1)
	{
		m_pNuiSensor_1->NuiShutdown();
	}

    if (m_hNextColorFrameEvent != INVALID_HANDLE_VALUE)
    {
        CloseHandle(m_hNextColorFrameEvent);
    }

	if (m_hNextColorFrameEvent_1 != INVALID_HANDLE_VALUE)
	{
		CloseHandle(m_hNextColorFrameEvent_1);
	}

    // clean up Direct2D renderer
    delete m_pDrawColor;
    m_pDrawColor = NULL;

	delete[] m_pTempColorBuffer;
	m_pTempColorBuffer = NULL;

	delete m_pDrawColor_1;
	m_pDrawColor_1 = NULL;

	delete[] m_pTempColorBuffer_1;
	m_pTempColorBuffer_1 = NULL;

    // clean up Direct2D
    SafeRelease(m_pD2DFactory);
    SafeRelease(m_pNuiSensor);

	SafeRelease(m_pD2DFactory_1);
	SafeRelease(m_pNuiSensor_1);
}

/// <summary>
/// Creates the main window and begins processing
/// </summary>
/// <param name="hInstance">handle to the application instance</param>
/// <param name="nCmdShow">whether to display minimized, maximized, or normally</param>
int CColorBasics::Run(HINSTANCE hInstance, int nCmdShow)
{
    MSG       msg = {0};
    WNDCLASS  wc;

    // Dialog custom window class
    ZeroMemory(&wc, sizeof(wc));
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.cbWndExtra    = DLGWINDOWEXTRA;
    wc.hInstance     = hInstance;
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hIcon         = LoadIconW(hInstance, MAKEINTRESOURCE(IDI_APP));
    wc.lpfnWndProc   = DefDlgProcW;
    wc.lpszClassName = L"ColorBasicsAppDlgWndClass";

    if (!RegisterClassW(&wc))
    {
        return 0;
    }

    // Create main application window
    HWND hWndApp = CreateDialogParamW(
        hInstance,
        MAKEINTRESOURCE(IDD_APP),
        NULL,
        (DLGPROC)CColorBasics::MessageRouter, 
        reinterpret_cast<LPARAM>(this));

    // Show window
    ShowWindow(hWndApp, nCmdShow);

    const int eventCount = 1;
    HANDLE hEvents[eventCount];

    // Main message loop
    while (WM_QUIT != msg.message)
    {
        hEvents[0] = m_hNextColorFrameEvent;

        // Check to see if we have either a message (by passing in QS_ALLINPUT)
        // Or a Kinect event (hEvents)
        // Update() will check for Kinect events individually, in case more than one are signalled
        MsgWaitForMultipleObjects(eventCount, hEvents, FALSE, INFINITE, QS_ALLINPUT);

        // Explicitly check the Kinect frame event since MsgWaitForMultipleObjects
        // can return for other reasons even though it is signaled.
        Update();

        while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE))
        {
            // If a dialog message will be taken care of by the dialog proc
            if ((hWndApp != NULL) && IsDialogMessageW(hWndApp, &msg))
            {
                continue;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    return static_cast<int>(msg.wParam);
}

/// <summary>
/// Main processing function
/// </summary>
void CColorBasics::Update()
{
    if (NULL == m_pNuiSensor)
    {
        return;
    }

	//if (NULL == m_pNuiSensor_1)
	//{
	//	return;
	//}


    if ( WAIT_OBJECT_0 == WaitForSingleObject(m_hNextColorFrameEvent, 0) )
    {
	    ProcessColor();
    }
}

/// <summary>
/// Handles window messages, passes most to the class instance to handle
/// </summary>
/// <param name="hWnd">window message is for</param>
/// <param name="uMsg">message</param>
/// <param name="wParam">message data</param>
/// <param name="lParam">additional message data</param>
/// <returns>result of message processing</returns>
LRESULT CALLBACK CColorBasics::MessageRouter(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CColorBasics* pThis = NULL;
    
    if (WM_INITDIALOG == uMsg)
    {
        pThis = reinterpret_cast<CColorBasics*>(lParam);
        SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
    }
    else
    {
        pThis = reinterpret_cast<CColorBasics*>(::GetWindowLongPtr(hWnd, GWLP_USERDATA));
    }

    if (pThis)
    {
        return pThis->DlgProc(hWnd, uMsg, wParam, lParam);
    }

    return 0;
}

/// <summary>
/// Handle windows messages for the class instance
/// </summary>
/// <param name="hWnd">window message is for</param>
/// <param name="uMsg">message</param>
/// <param name="wParam">message data</param>
/// <param name="lParam">additional message data</param>
/// <returns>result of message processing</returns>
LRESULT CALLBACK CColorBasics::DlgProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_INITDIALOG:
        {
            // Bind application window handle
            m_hWnd = hWnd;

            // Init Direct2D
            D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &m_pD2DFactory);
			D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &m_pD2DFactory_1);
            // Create and initialize a new Direct2D image renderer (take a look at ImageRenderer.h)
            // We'll use this to draw the data we receive from the Kinect to the screen
            m_pDrawColor = new ImageRenderer();
			m_pDrawColor_1 = new ImageRenderer();
            HRESULT hr = m_pDrawColor->Initialize(GetDlgItem(m_hWnd, IDC_VIDEOVIEW), m_pD2DFactory, cColorWidth, cColorHeight, cColorWidth * sizeof(long));
			HRESULT hr1 = m_pDrawColor_1->Initialize(GetDlgItem(m_hWnd, IDC_VIDEOVIEW1), m_pD2DFactory_1, cColorWidth, cColorHeight, cColorWidth * sizeof(long));
            if (FAILED(hr)&&FAILED(hr1))
            {
                SetStatusMessage(L"Failed to initialize the Direct2D draw device.");
            }

            // Look for a connected Kinect, and create it if found
      //      CreateFirstConnected();
        }
        break;

        // If the titlebar X is clicked, destroy app
        case WM_CLOSE:
            DestroyWindow(hWnd);
            break;

        case WM_DESTROY:
            // Quit the main message pump
            PostQuitMessage(0);
            break;

        // Handle button press
        case WM_COMMAND:
            // If it was for the screenshot control and a button clicked event, save a screenshot next frame 
            if (IDC_BUTTON_SCREENSHOT == LOWORD(wParam) && BN_CLICKED == HIWORD(wParam))
            {
                m_bSaveScreenshot = true;
				m_bSaveScreenshot_1 = true;

				ChooseNumberImage		nParam;

				INT_PTR result = DialogBoxParam(GetModuleHandle(NULL),
					MAKEINTRESOURCE(IDD_NUM_IMAGE), hWnd,
					NumberDeviceDlgProc, (LPARAM)&nParam);

				n = nParam.selection;
            }

			if (IDC_CHECK1 == LOWORD(wParam) && BN_CLICKED == HIWORD(wParam))
			{
				if (SendDlgItemMessage(hWnd, IDC_CHECK1, BM_GETCHECK, 0, 0))
					m_bOffIrEmitter = true;
				else
					m_bOffIrEmitter = false;
			}

			if (IDC_CHECK2 == LOWORD(wParam) && BN_CLICKED == HIWORD(wParam))
			{
				if (SendDlgItemMessage(hWnd, IDC_CHECK2, BM_GETCHECK, 0, 0))
					m_continuouscap = true;
				else
					m_continuouscap = false;
			}


			if (IDC_StartScreen == LOWORD(wParam) && BN_CLICKED == HIWORD(wParam))
			{
				CreateFirstConnected();
			}
			break;
    }

    return FALSE;
}

/// <summary>
/// Create the first connected Kinect found 
/// </summary>
/// <returns>indicates success or failure</returns>
HRESULT CColorBasics::CreateFirstConnected()
{
    INuiSensor * pNuiSensor;
	INuiSensor * pNuiSensor_1;
    HRESULT hr;

    int iSensorCount = 0;
    hr = NuiGetSensorCount(&iSensorCount);
    if (FAILED(hr))
    {
        return hr;
    }

    // First Sensor

        // Create the sensor so we can check status, if we can't create it, move on to the next
        hr = NuiCreateSensorByIndex(0, &pNuiSensor);
        // Get the status of the sensor, and if connected, then we can initialize it
        hr = pNuiSensor->NuiStatus();
        if (S_OK == hr)
        {
            m_pNuiSensor = pNuiSensor;
        }
        // This sensor wasn't OK, so release it since we're not using it
		if (S_OK != hr)
		{
			pNuiSensor->Release();
		}
	
	// Second Sensor

			hr = NuiCreateSensorByIndex(1, &pNuiSensor_1);
			// Get the status of the sensor, and if connected, then we can initialize it
			hr = pNuiSensor_1->NuiStatus();
			if (S_OK == hr)
			{
				m_pNuiSensor_1 = pNuiSensor_1;
			}
			// This sensor wasn't OK, so release it since we're not using it
			if (S_OK != hr)
			{
				pNuiSensor_1->Release();
			}


    if (NULL != m_pNuiSensor)
    {
        // Initialize the Kinect and specify that we'll be using color
        hr = m_pNuiSensor->NuiInitialize(NUI_INITIALIZE_FLAG_USES_COLOR); 
        if (SUCCEEDED(hr))
        {
            // Create an event that will be signaled when color data is available
            m_hNextColorFrameEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

			if (m_bOffIrEmitter) {
				m_pNuiSensor->NuiSetForceInfraredEmitterOff(1);			//0 to turn it on, 1 to turn if off. We need to turn of for Camera Calibration
				SetStatusMessage(L"IR Emitter is Off");
			}
			else if (m_bOffIrEmitter == false) {
				m_pNuiSensor->NuiSetForceInfraredEmitterOff(0);
				 SetStatusMessage(L"IR Emitteris On");
			}

            // Open a color image stream to receive color frames
            hr = m_pNuiSensor->NuiImageStreamOpen(
				NUI_IMAGE_TYPE_COLOR_INFRARED,				
                NUI_IMAGE_RESOLUTION_640x480,
                0,
                2,
                m_hNextColorFrameEvent,
                &m_pColorStreamHandle);
        }
    }


	if (NULL != m_pNuiSensor_1)
	{
		// Initialize the Kinect and specify that we'll be using color
		hr = m_pNuiSensor_1->NuiInitialize(NUI_INITIALIZE_FLAG_USES_COLOR);
		if (SUCCEEDED(hr))
		{
			// Create an event that will be signaled when color data is available
			m_hNextColorFrameEvent_1 = CreateEvent(NULL, TRUE, FALSE, NULL);

			if (m_bOffIrEmitter) {
				m_pNuiSensor_1->NuiSetForceInfraredEmitterOff(1);			//0 to turn it on, 1 to turn if off. We need to turn of for Camera Calibration
				SetStatusMessage(L"IR Emitter is Off");
			}
			else if (m_bOffIrEmitter == false) {
				m_pNuiSensor_1->NuiSetForceInfraredEmitterOff(0);
				SetStatusMessage(L"IR Emitteris On");
			}

			// Open a color image stream to receive color frames
			hr = m_pNuiSensor_1->NuiImageStreamOpen(
				NUI_IMAGE_TYPE_COLOR_INFRARED,
				NUI_IMAGE_RESOLUTION_640x480,
				0,
				2,
				m_hNextColorFrameEvent_1,
				&m_pColorStreamHandle_1);
		}
	}

    if (NULL == m_pNuiSensor || FAILED(hr))
    {
        SetStatusMessage(L"No ready Kinect found!");
        return E_FAIL;
    }

    return hr;
}

/// <summary>
/// Get the name of the file where screenshot will be stored.
/// </summary>
/// <param name="screenshotName">
/// [out] String buffer that will receive screenshot file name.
/// </param>
/// <param name="screenshotNameSize">
/// [in] Number of characters in screenshotName string buffer.
/// </param>
/// <returns>
/// S_OK on success, otherwise failure code.
/// </returns>
//HRESULT GetScreenshotFileName(wchar_t *screenshotName, UINT screenshotNameSize, int i, int n)
//{
//    wchar_t *knownPath = NULL;
//    HRESULT hr = SHGetKnownFolderPath(FOLDERID_Pictures, 0, NULL, &knownPath);
//
//        // Get the time
//        //wchar_t timeString[MAX_PATH];
//        //GetTimeFormatEx(NULL, 0, NULL, L"hh'-'mm'-'ss", timeString, _countof(timeString));
//
//		if (i == 0) 
//			StringCchPrintfW(screenshotName, screenshotNameSize, L"%s_Cam01_%02d.bmp", ofn.lpstrFile, n);	
//		else if (i == 1)
//			StringCchPrintfW(screenshotName, screenshotNameSize, L"%s_Cam02_%02d.bmp", ofn.lpstrFile, n);
// 
//
//    CoTaskMemFree(knownPath);
//    return hr;
//}

/// <summary>
/// Handle new color data
/// </summary>
/// <returns>indicates success or failure</returns>
void CColorBasics::ProcessColor()
{
    HRESULT hr;
    NUI_IMAGE_FRAME imageFrame;	

    // Attempt to get the color frame
    hr = m_pNuiSensor->NuiImageStreamGetNextFrame(m_pColorStreamHandle, 0, &imageFrame);
    if (FAILED(hr))
    {
        return;
    }

    INuiFrameTexture * pTexture = imageFrame.pFrameTexture;
    NUI_LOCKED_RECT LockedRect;

    // Lock the frame data so the Kinect knows not to modify it while we're reading it
    pTexture->LockRect(0, &LockedRect, NULL, 0);	

    // Make sure we've received valid data
    if (LockedRect.Pitch != 0)
    {
		//// This code was copied from the InfaredBasics-D2D project
        // Draw the IR data with Direct2D
		if (m_pTempColorBuffer == NULL)
		{
			m_pTempColorBuffer = new RGBQUAD[cColorWidth * cColorHeight];
		}

		for (int i = 0; i < cColorWidth * cColorHeight; ++i)
		{
			BYTE intensity = reinterpret_cast<USHORT*>(LockedRect.pBits)[i] >> 8;

			RGBQUAD *pQuad = &m_pTempColorBuffer[i];
			pQuad->rgbBlue = intensity;
			pQuad->rgbGreen = intensity;
			pQuad->rgbRed = intensity;
			pQuad->rgbReserved = 255;
		}

		// Draw the data with Direct2D
		m_pDrawColor->Draw(reinterpret_cast<BYTE*>(m_pTempColorBuffer), cColorWidth * cColorHeight * sizeof(RGBQUAD));
		////

        // If the user pressed the screenshot button, save a screenshot
        if (m_bSaveScreenshot)
        {

            //WCHAR statusMessage[cStatusMessageMaxLen];

            // Retrieve the path to My Photos
            WCHAR screenshotPath[MAX_PATH];

			StringCchPrintfW(screenshotPath, _countof(screenshotPath), L"%sCam01_%03d.bmp", m_str, m);	//Change 02d to 03d to save hundred of images

            // Write out the bitmap to disk
            hr = SaveBitmapToFile(reinterpret_cast<BYTE *>(m_pTempColorBuffer), cColorWidth, cColorHeight, 32, screenshotPath);

			//Stop printing out to dialog in order to save time
            //if (SUCCEEDED(hr))
            //{
            //    // Set the status bar to show where the screenshot was saved
            //    StringCchPrintf( statusMessage, cStatusMessageMaxLen, L"Kinect Sensor 1 saved to %s", screenshotPath);
            //}
            //else
            //{
            //    StringCchPrintf( statusMessage, cStatusMessageMaxLen, L"1 Failed to write screenshot to %s", screenshotPath);
            //}
            //SetStatusMessage(statusMessage);
        }
    }

    // We're done with the texture so unlock it
    pTexture->UnlockRect(0);

    // Release the frame
    m_pNuiSensor->NuiImageStreamReleaseFrame(m_pColorStreamHandle, &imageFrame);

	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// 2nd IR camera
	
	NUI_IMAGE_FRAME imageFrame_1;

	// Attempt to get the color frame
	hr = m_pNuiSensor_1->NuiImageStreamGetNextFrame(m_pColorStreamHandle_1, 0, &imageFrame_1);
	if (FAILED(hr))
	{
		return;
	}

	INuiFrameTexture * pTexture1 = imageFrame_1.pFrameTexture;
//	NUI_LOCKED_RECT LockedRect;

	// Lock the frame data so the Kinect knows not to modify it while we're reading it
	pTexture1->LockRect(0, &LockedRect, NULL, 0);

	// Make sure we've received valid data
	if (LockedRect.Pitch != 0)
	{
		//// This code was copied from the InfaredBasics-D2D project
		// Draw the IR data with Direct2D
		if (m_pTempColorBuffer_1 == NULL)
		{
			m_pTempColorBuffer_1 = new RGBQUAD[cColorWidth * cColorHeight];
		}

		for (int i = 0; i < cColorWidth * cColorHeight; ++i)
		{
			BYTE intensity = reinterpret_cast<USHORT*>(LockedRect.pBits)[i] >> 8;

			RGBQUAD *pQuad = &m_pTempColorBuffer_1[i];
			pQuad->rgbBlue = intensity;
			pQuad->rgbGreen = intensity;
			pQuad->rgbRed = intensity;
			pQuad->rgbReserved = 255;
		}

		// Draw the data with Direct2D
		m_pDrawColor_1->Draw(reinterpret_cast<BYTE*>(m_pTempColorBuffer_1), cColorWidth * cColorHeight * sizeof(RGBQUAD));
		////

		// If the user pressed the screenshot button, save a screenshot
		if (m_bSaveScreenshot_1)
		{
//			WCHAR statusMessage[cStatusMessageMaxLen];

			// Retrieve the path to My Photos
			WCHAR screenshotPath[MAX_PATH];
			StringCchPrintfW(screenshotPath, _countof(screenshotPath), L"%sCam02_%03d.bmp", m_str, m);	//Change 02d to 03d to save hundred of images

			// Write out the bitmap to disk
			hr = SaveBitmapToFile(reinterpret_cast<BYTE *>(m_pTempColorBuffer_1), cColorWidth, cColorHeight, 32, screenshotPath);

			//Stop printing out to dialog in order to save time
			//if (SUCCEEDED(hr))
			//{
			//	// Set the status bar to show where the screenshot was saved
			//	StringCchPrintf(statusMessage, cStatusMessageMaxLen, L"Kinect Sensor 2 saved to %s", screenshotPath);
			//}
			//else
			//{
			//	StringCchPrintf(statusMessage, cStatusMessageMaxLen, L"2 Failed to write screenshot to %s", screenshotPath);
			//}
			//SetStatusMessage(statusMessage);

			if (m_continuouscap == false)
			{
				std::string test = std::to_string(m);
				MessageBox(NULL, LPCWSTR(test.c_str()), L"Result", MB_OK);
			}
			m++;

			if (m == n) {
				m_bSaveScreenshot = false;
				m_bSaveScreenshot_1 = false;
				m = 0;
				MessageBox(NULL, L"Capture is done!!!", L"Notification", MB_OK);
			}
		}
	}

	// We're done with the texture so unlock it
	pTexture1->UnlockRect(0);

	// Release the frame
	m_pNuiSensor_1->NuiImageStreamReleaseFrame(m_pColorStreamHandle_1, &imageFrame_1);
}

/// <summary>
/// Set the status bar message
/// </summary>
/// <param name="szMessage">message to display</param>
void CColorBasics::SetStatusMessage(WCHAR * szMessage)
{
    SendDlgItemMessageW(m_hWnd, IDC_STATUS, WM_SETTEXT, 0, (LPARAM)szMessage);
}

/// <summary>
/// Save passed in image data to disk as a bitmap
/// </summary>
/// <param name="pBitmapBits">image data to save</param>
/// <param name="lWidth">width (in pixels) of input image data</param>
/// <param name="lHeight">height (in pixels) of input image data</param>
/// <param name="wBitsPerPixel">bits per pixel of image data</param>
/// <param name="lpszFilePath">full file path to output bitmap to</param>
/// <returns>indicates success or failure</returns>
HRESULT CColorBasics::SaveBitmapToFile(BYTE* pBitmapBits, LONG lWidth, LONG lHeight, WORD wBitsPerPixel, LPCWSTR lpszFilePath)
{
    DWORD dwByteCount = lWidth * lHeight * (wBitsPerPixel / 8);

    BITMAPINFOHEADER bmpInfoHeader = {0};

    bmpInfoHeader.biSize        = sizeof(BITMAPINFOHEADER);  // Size of the header
    bmpInfoHeader.biBitCount    = wBitsPerPixel;             // Bit count
    bmpInfoHeader.biCompression = BI_RGB;                    // Standard RGB, no compression
    bmpInfoHeader.biWidth       = lWidth;                    // Width in pixels
    bmpInfoHeader.biHeight      = -lHeight;                  // Height in pixels, negative indicates it's stored right-side-up
    bmpInfoHeader.biPlanes      = 1;                         // Default
    bmpInfoHeader.biSizeImage   = dwByteCount;               // Image size in bytes

    BITMAPFILEHEADER bfh = {0};

    bfh.bfType    = 0x4D42;                                           // 'M''B', indicates bitmap
    bfh.bfOffBits = bmpInfoHeader.biSize + sizeof(BITMAPFILEHEADER);  // Offset to the start of pixel data
    bfh.bfSize    = bfh.bfOffBits + bmpInfoHeader.biSizeImage;        // Size of image + headers

    // Create the file on disk to write to
    HANDLE hFile = CreateFileW(lpszFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    // Return if error opening file
    if (NULL == hFile) 
    {
        return E_ACCESSDENIED;
    }

    DWORD dwBytesWritten = 0;
    
    // Write the bitmap file header
    if ( !WriteFile(hFile, &bfh, sizeof(bfh), &dwBytesWritten, NULL) )
    {
        CloseHandle(hFile);
        return E_FAIL;
    }
    
    // Write the bitmap info header
    if ( !WriteFile(hFile, &bmpInfoHeader, sizeof(bmpInfoHeader), &dwBytesWritten, NULL) )
    {
        CloseHandle(hFile);
        return E_FAIL;
    }
    
    // Write the RGB Data
    if ( !WriteFile(hFile, pBitmapBits, bmpInfoHeader.biSizeImage, &dwBytesWritten, NULL) )
    {
        CloseHandle(hFile);
        return E_FAIL;
    }    

    // Close the file
    CloseHandle(hFile);
    return S_OK;
}
