//------------------------------------------------------------------------------
// <copyright file="ColorBasics.h" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

#pragma once

#include "resource.h"
#include "NuiApi.h"
#include "ImageRenderer.h"


struct ChooseNumberImage
{
	UINT32 count;
	UINT32 selection;
};

class CColorBasics
{
    static const int        cColorWidth  = 640;
    static const int        cColorHeight = 480;

    static const int        cStatusMessageMaxLen = MAX_PATH*2;

public:
    /// <summary>
    /// Constructor
    /// </summary>
    CColorBasics();

    /// <summary>
    /// Destructor
    /// </summary>
    ~CColorBasics();

    /// <summary>
    /// Handles window messages, passes most to the class instance to handle
    /// </summary>
    /// <param name="hWnd">window message is for</param>
    /// <param name="uMsg">message</param>
    /// <param name="wParam">message data</param>
    /// <param name="lParam">additional message data</param>
    /// <returns>result of message processing</returns>
    static LRESULT CALLBACK MessageRouter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);

    /// <summary>
    /// Handle windows messages for a class instance
    /// </summary>
    /// <param name="hWnd">window message is for</param>
    /// <param name="uMsg">message</param>
    /// <param name="wParam">message data</param>
    /// <param name="lParam">additional message data</param>
    /// <returns>result of message processing</returns>
    LRESULT CALLBACK        DlgProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);

    /// <summary>
    /// Creates the main window and begins processing
    /// </summary>
    /// <param name="hInstance"></param>
    /// <param name="nCmdShow"></param>
    int                     Run(HINSTANCE hInstance, int nCmdShow);

private:
    HWND                    m_hWnd;

    bool                    m_bSaveScreenshot, m_bSaveScreenshot_1;
	bool					m_bOffIrEmitter;
	bool					m_continuouscap;
	int						n;
	int						m = 0;

    // Current Kinect
    INuiSensor*             m_pNuiSensor;
	INuiSensor*             m_pNuiSensor_1;

    // Direct2D
    ImageRenderer*          m_pDrawColor;
    ID2D1Factory*           m_pD2DFactory;

	ImageRenderer*          m_pDrawColor_1;
	ID2D1Factory*           m_pD2DFactory_1;
    
    HANDLE                  m_pColorStreamHandle;
    HANDLE                  m_hNextColorFrameEvent;

	RGBQUAD*                m_pTempColorBuffer;

	HANDLE                  m_pColorStreamHandle_1;
	HANDLE                  m_hNextColorFrameEvent_1;

	RGBQUAD*                m_pTempColorBuffer_1;
	wchar_t					*m_str = L"C:/Temp/";
	 


    /// <summary>
    /// Main processing function
    /// </summary>
    void                    Update();

    /// <summary>
    /// Create the first connected Kinect found 
    /// </summary>
    /// <returns>S_OK on success, otherwise failure code</returns>
    HRESULT                 CreateFirstConnected();

    /// <summary>
    /// Handle new color data
    /// </summary>
    void                    ProcessColor();

    /// <summary>
    /// Set the status bar message
    /// </summary>
    /// <param name="szMessage">message to display</param>
    void                    SetStatusMessage(WCHAR* szMessage);

    /// <summary>
    /// Save passed in image data to disk as a bitmap
    /// </summary>
    /// <param name="pBitmapBits">image data to save</param>
    /// <param name="lWidth">width (in pixels) of input image data</param>
    /// <param name="lHeight">height (in pixels) of input image data</param>
    /// <param name="wBitsPerPixel">bits per pixel of image data</param>
    /// <param name="lpszFilePath">full file path to output bitmap to</param>
    /// <returns>S_OK on success, otherwise failure code</returns>
    HRESULT                 SaveBitmapToFile(BYTE* pBitmapBits, LONG lWidth, LONG lHeight, WORD wBitsPerPixel, LPCTSTR lpszFilePath);
	HRESULT					setLaser(bool on);
};
