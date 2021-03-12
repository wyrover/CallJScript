// CallJScript.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "SimpleScriptSite.h"
#include <atlbase.h>
#include <activscp.h>
#include <atlwin.h>
#include <iostream>
#include <stdio.h>
#include <vector>
#include <fstream>


void testExpression(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
	wprintf(L"%s: %s: ", prefix, script);
	HRESULT hr = S_OK;
	CComVariant result;
	EXCEPINFO ei = { };
	hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &result, &ei);
	CComVariant resultBSTR;
	VariantChangeType(&resultBSTR, &result, 0, VT_BSTR);
	wprintf(L"%s\n", V_BSTR(&resultBSTR));
}

void testScript(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
	wprintf(L"%s: %s\n", prefix, script);
	HRESULT hr = S_OK;
	CComVariant result;
	EXCEPINFO ei = { };
	hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, 0, &result, &ei);
}


BOOL get_file_contents(const char* filename, std::vector<BYTE>& out_buffer) {
	std::ifstream in(filename, std::ios::in | std::ios::binary);

	if (in) {
		in.seekg(0, std::ios::end);
		out_buffer.resize((size_t)in.tellg()+1, 0);
		in.seekg(0, std::ios::beg);
		in.read((char*)&out_buffer[0], out_buffer.size());
		in.close();
		return TRUE;
	}

	return FALSE;
}



std::wstring convert_to(const char* const& from)
{
	int len = static_cast<int>(strlen(from) + 1);
	std::vector<wchar_t> wideBuf(len);
	int ret = MultiByteToWideChar(CP_ACP, 0, from, len, &wideBuf[0], len);

	if (ret == 0)
		return L"";
	else
		return std::wstring(&wideBuf[0]);
}



int _tmain(int argc, _TCHAR* argv[])
{
	std::vector<BYTE> buffer;

	if (!get_file_contents(".\\aes.js", buffer)) {
		return 1;
	}


	std::wstring content = convert_to((const char*)&buffer[0]);



	HRESULT hr = S_OK;
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	// Initialize
	CSimpleScriptSite* pScriptSite = new CSimpleScriptSite();
	
	//CComVariant result;
	//EXCEPINFO ei = { };

	//hr = pScriptSite->Eval(content.c_str(), &result);
	//if (!SUCCEEDED(hr)) {
	//	printf("failed!\n");
	//}

	//
	//CComVariant resultBSTR;
	//VariantChangeType(&resultBSTR, &result, 0, VT_BSTR);
	//wprintf(L"%s\n", V_BSTR(&resultBSTR));
	//

	pScriptSite->AddScript(content.c_str());

	BSTR arg1 = ::SysAllocString(L"123456");
	
	VARIANT vRes;
	DISPPARAMS args;
	args.cNamedArgs = 0;
	args.cArgs = 1;
	VARIANTARG* pVariant = new VARIANTARG;
	args.rgvarg = pVariant;
	args.rgvarg[0].vt = VT_BSTR;
	args.rgvarg[0].bstrVal = arg1;
	pScriptSite->Run((WCHAR*)L"encryptAES", &args, &vRes);

	wprintf(L"retval: %s\n", vRes.bstrVal);

	SysFreeString(arg1);
	
	
	delete pVariant;

	


	


	
	pScriptSite->Release();
	pScriptSite = NULL;

	::CoUninitialize();
	return 0;
}