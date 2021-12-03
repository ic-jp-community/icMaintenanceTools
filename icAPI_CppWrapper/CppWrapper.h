// CppWrapper.h : CppWrapper.DLL のメイン ヘッダー ファイル
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH に対してこのファイルをインクルードする前に 'stdafx.h' をインクルードしてください"
#endif

#include "resource.h"		// メイン シンボル


// CCppWrapperApp
// このクラスの実装に関しては CppWrapper.cpp を参照してください。
//

class CCppWrapperApp : public CWinApp
{
public:
	CCppWrapperApp();

// オーバーライド
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};
extern "C" {
	__declspec(dllexport) CWnd* HwndToCwnd(HWND* hwnd);
}
