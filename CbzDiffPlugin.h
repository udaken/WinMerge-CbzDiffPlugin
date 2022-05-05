

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0626 */
/* at Tue Jan 19 12:14:07 2038
 */
/* Compiler settings for CbzDiffPlugin.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0626 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __CbzDiffPlugin_h__
#define __CbzDiffPlugin_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if _CONTROL_FLOW_GUARD_XFG
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __IWinMergeScript_FWD_DEFINED__
#define __IWinMergeScript_FWD_DEFINED__
typedef interface IWinMergeScript IWinMergeScript;

#endif 	/* __IWinMergeScript_FWD_DEFINED__ */


#ifndef __WinMergeScript_FWD_DEFINED__
#define __WinMergeScript_FWD_DEFINED__

#ifdef __cplusplus
typedef class WinMergeScript WinMergeScript;
#else
typedef struct WinMergeScript WinMergeScript;
#endif /* __cplusplus */

#endif 	/* __WinMergeScript_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IWinMergeScript_INTERFACE_DEFINED__
#define __IWinMergeScript_INTERFACE_DEFINED__

/* interface IWinMergeScript */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IWinMergeScript;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("7C37C1C6-FD13-4285-857D-35FAED278EC3")
    IWinMergeScript : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginEvent( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginDescription( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginFileFilters( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginIsAutomatic( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginExtendedProperties( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_PluginUnpackedFileExtension( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UnpackBufferA( 
            /* [in] */ SAFEARRAY * *pBuffer,
            /* [in] */ INT *pSize,
            /* [in] */ VARIANT_BOOL *pbChanged,
            /* [in] */ INT *pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PackBufferA( 
            /* [in] */ SAFEARRAY * *pBuffer,
            /* [in] */ INT *pSize,
            /* [in] */ VARIANT_BOOL *pbChanged,
            /* [in] */ INT subcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE UnpackFile( 
            /* [in] */ BSTR fileSrc,
            /* [in] */ BSTR fileDst,
            VARIANT_BOOL *pbChanged,
            INT *pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PackFile( 
            /* [in] */ BSTR fileSrc,
            /* [in] */ BSTR fileDst,
            VARIANT_BOOL *pbChanged,
            INT pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ShowSettingsDialog( 
            /* [retval][out] */ VARIANT_BOOL *pbHandled) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IWinMergeScriptVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IWinMergeScript * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IWinMergeScript * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IWinMergeScript * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IWinMergeScript * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IWinMergeScript * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IWinMergeScript * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IWinMergeScript * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginEvent)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginEvent )( 
            IWinMergeScript * This,
            /* [retval][out] */ BSTR *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginDescription)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginDescription )( 
            IWinMergeScript * This,
            /* [retval][out] */ BSTR *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginFileFilters)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginFileFilters )( 
            IWinMergeScript * This,
            /* [retval][out] */ BSTR *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginIsAutomatic)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginIsAutomatic )( 
            IWinMergeScript * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginExtendedProperties)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginExtendedProperties )( 
            IWinMergeScript * This,
            /* [retval][out] */ BSTR *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, get_PluginUnpackedFileExtension)
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_PluginUnpackedFileExtension )( 
            IWinMergeScript * This,
            /* [retval][out] */ BSTR *pVal);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, UnpackBufferA)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *UnpackBufferA )( 
            IWinMergeScript * This,
            /* [in] */ SAFEARRAY * *pBuffer,
            /* [in] */ INT *pSize,
            /* [in] */ VARIANT_BOOL *pbChanged,
            /* [in] */ INT *pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, PackBufferA)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *PackBufferA )( 
            IWinMergeScript * This,
            /* [in] */ SAFEARRAY * *pBuffer,
            /* [in] */ INT *pSize,
            /* [in] */ VARIANT_BOOL *pbChanged,
            /* [in] */ INT subcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, UnpackFile)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *UnpackFile )( 
            IWinMergeScript * This,
            /* [in] */ BSTR fileSrc,
            /* [in] */ BSTR fileDst,
            VARIANT_BOOL *pbChanged,
            INT *pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, PackFile)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *PackFile )( 
            IWinMergeScript * This,
            /* [in] */ BSTR fileSrc,
            /* [in] */ BSTR fileDst,
            VARIANT_BOOL *pbChanged,
            INT pSubcode,
            /* [retval][out] */ VARIANT_BOOL *pbSuccess);
        
        DECLSPEC_XFGVIRT(IWinMergeScript, ShowSettingsDialog)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ShowSettingsDialog )( 
            IWinMergeScript * This,
            /* [retval][out] */ VARIANT_BOOL *pbHandled);
        
        END_INTERFACE
    } IWinMergeScriptVtbl;

    interface IWinMergeScript
    {
        CONST_VTBL struct IWinMergeScriptVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWinMergeScript_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IWinMergeScript_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IWinMergeScript_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IWinMergeScript_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IWinMergeScript_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IWinMergeScript_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IWinMergeScript_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IWinMergeScript_get_PluginEvent(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginEvent(This,pVal) ) 

#define IWinMergeScript_get_PluginDescription(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginDescription(This,pVal) ) 

#define IWinMergeScript_get_PluginFileFilters(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginFileFilters(This,pVal) ) 

#define IWinMergeScript_get_PluginIsAutomatic(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginIsAutomatic(This,pVal) ) 

#define IWinMergeScript_get_PluginExtendedProperties(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginExtendedProperties(This,pVal) ) 

#define IWinMergeScript_get_PluginUnpackedFileExtension(This,pVal)	\
    ( (This)->lpVtbl -> get_PluginUnpackedFileExtension(This,pVal) ) 

#define IWinMergeScript_UnpackBufferA(This,pBuffer,pSize,pbChanged,pSubcode,pbSuccess)	\
    ( (This)->lpVtbl -> UnpackBufferA(This,pBuffer,pSize,pbChanged,pSubcode,pbSuccess) ) 

#define IWinMergeScript_PackBufferA(This,pBuffer,pSize,pbChanged,subcode,pbSuccess)	\
    ( (This)->lpVtbl -> PackBufferA(This,pBuffer,pSize,pbChanged,subcode,pbSuccess) ) 

#define IWinMergeScript_UnpackFile(This,fileSrc,fileDst,pbChanged,pSubcode,pbSuccess)	\
    ( (This)->lpVtbl -> UnpackFile(This,fileSrc,fileDst,pbChanged,pSubcode,pbSuccess) ) 

#define IWinMergeScript_PackFile(This,fileSrc,fileDst,pbChanged,pSubcode,pbSuccess)	\
    ( (This)->lpVtbl -> PackFile(This,fileSrc,fileDst,pbChanged,pSubcode,pbSuccess) ) 

#define IWinMergeScript_ShowSettingsDialog(This,pbHandled)	\
    ( (This)->lpVtbl -> ShowSettingsDialog(This,pbHandled) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IWinMergeScript_INTERFACE_DEFINED__ */



#ifndef __CbzDiffPlugin_LIBRARY_DEFINED__
#define __CbzDiffPlugin_LIBRARY_DEFINED__

/* library CbzDiffPlugin */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_CbzDiffPlugin;

EXTERN_C const CLSID CLSID_WinMergeScript;

#ifdef __cplusplus

class DECLSPEC_UUID("2980B00C-CF8C-405D-8FC9-338FCBCF0B48")
WinMergeScript;
#endif
#endif /* __CbzDiffPlugin_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long *, LPSAFEARRAY * ); 

unsigned long             __RPC_USER  BSTR_UserSize64(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal64(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal64(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree64(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize64(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal64(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal64(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree64(     unsigned long *, LPSAFEARRAY * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


