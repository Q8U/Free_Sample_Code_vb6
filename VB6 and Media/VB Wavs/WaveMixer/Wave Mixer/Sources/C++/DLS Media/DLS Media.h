/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Apr 25 07:50:41 2002
 */
/* Compiler settings for C:\Data\Sources\C++\DLS Media\DLS Media.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __DLS_Media_h__
#define __DLS_Media_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IWaveMixer_FWD_DEFINED__
#define __IWaveMixer_FWD_DEFINED__
typedef interface IWaveMixer IWaveMixer;
#endif 	/* __IWaveMixer_FWD_DEFINED__ */


#ifndef ___IWaveMixerEvents_FWD_DEFINED__
#define ___IWaveMixerEvents_FWD_DEFINED__
typedef interface _IWaveMixerEvents _IWaveMixerEvents;
#endif 	/* ___IWaveMixerEvents_FWD_DEFINED__ */


#ifndef __WaveMixer_FWD_DEFINED__
#define __WaveMixer_FWD_DEFINED__

#ifdef __cplusplus
typedef class WaveMixer WaveMixer;
#else
typedef struct WaveMixer WaveMixer;
#endif /* __cplusplus */

#endif 	/* __WaveMixer_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IWaveMixer_INTERFACE_DEFINED__
#define __IWaveMixer_INTERFACE_DEFINED__

/* interface IWaveMixer */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IWaveMixer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1D847139-91A6-4395-9414-AC524F37A264")
    IWaveMixer : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [in] */ BSTR bstrFileName,
            /* [retval][out] */ long __RPC_FAR *pnID) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LoadedWaves( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetSoundInfo( 
            /* [in] */ int nIndex,
            /* [out] */ long __RPC_FAR *pnID,
            /* [out] */ BSTR __RPC_FAR *pbstrFileName,
            /* [out] */ long __RPC_FAR *pnOriginalFrequency,
            /* [out] */ long __RPC_FAR *pnChannels,
            /* [out] */ long __RPC_FAR *pnBits,
            /* [out] */ long __RPC_FAR *pnSize) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Initialize( 
            /* [in] */ long hWnd) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Play( 
            /* [in] */ long nID,
            /* [in] */ BOOL bLoop) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ long nID) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetVolume( 
            /* [in] */ long nID,
            /* [in] */ long nVolume) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetPan( 
            /* [in] */ long nID,
            /* [in] */ long nPan) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Stop( 
            /* [in] */ long nID) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetFrequency( 
            /* [in] */ long nID,
            /* [in] */ long nFrequency) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ResetFrequency( 
            /* [in] */ long nID) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFrequency( 
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnFrequency) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetPan( 
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnPan) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetVolume( 
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnVolume) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_UsedMemory( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_MaxMemory( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_MaxMemory( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetName( 
            /* [in] */ long nID,
            /* [out] */ BSTR __RPC_FAR *pbstrName) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StopAll( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IWaveMixerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IWaveMixer __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IWaveMixer __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IWaveMixer __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Add )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ BSTR bstrFileName,
            /* [retval][out] */ long __RPC_FAR *pnID);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Clear )( 
            IWaveMixer __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_LoadedWaves )( 
            IWaveMixer __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetSoundInfo )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ int nIndex,
            /* [out] */ long __RPC_FAR *pnID,
            /* [out] */ BSTR __RPC_FAR *pbstrFileName,
            /* [out] */ long __RPC_FAR *pnOriginalFrequency,
            /* [out] */ long __RPC_FAR *pnChannels,
            /* [out] */ long __RPC_FAR *pnBits,
            /* [out] */ long __RPC_FAR *pnSize);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Initialize )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long hWnd);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Play )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [in] */ BOOL bLoop);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Remove )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetVolume )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [in] */ long nVolume);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetPan )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [in] */ long nPan);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Stop )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetFrequency )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [in] */ long nFrequency);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ResetFrequency )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFrequency )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnFrequency);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetPan )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnPan);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetVolume )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [out] */ long __RPC_FAR *pnVolume);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_UsedMemory )( 
            IWaveMixer __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_MaxMemory )( 
            IWaveMixer __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_MaxMemory )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetName )( 
            IWaveMixer __RPC_FAR * This,
            /* [in] */ long nID,
            /* [out] */ BSTR __RPC_FAR *pbstrName);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *StopAll )( 
            IWaveMixer __RPC_FAR * This);
        
        END_INTERFACE
    } IWaveMixerVtbl;

    interface IWaveMixer
    {
        CONST_VTBL struct IWaveMixerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IWaveMixer_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IWaveMixer_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IWaveMixer_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IWaveMixer_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IWaveMixer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IWaveMixer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IWaveMixer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IWaveMixer_Add(This,bstrFileName,pnID)	\
    (This)->lpVtbl -> Add(This,bstrFileName,pnID)

#define IWaveMixer_Clear(This)	\
    (This)->lpVtbl -> Clear(This)

#define IWaveMixer_get_LoadedWaves(This,pVal)	\
    (This)->lpVtbl -> get_LoadedWaves(This,pVal)

#define IWaveMixer_GetSoundInfo(This,nIndex,pnID,pbstrFileName,pnOriginalFrequency,pnChannels,pnBits,pnSize)	\
    (This)->lpVtbl -> GetSoundInfo(This,nIndex,pnID,pbstrFileName,pnOriginalFrequency,pnChannels,pnBits,pnSize)

#define IWaveMixer_Initialize(This,hWnd)	\
    (This)->lpVtbl -> Initialize(This,hWnd)

#define IWaveMixer_Play(This,nID,bLoop)	\
    (This)->lpVtbl -> Play(This,nID,bLoop)

#define IWaveMixer_Remove(This,nID)	\
    (This)->lpVtbl -> Remove(This,nID)

#define IWaveMixer_SetVolume(This,nID,nVolume)	\
    (This)->lpVtbl -> SetVolume(This,nID,nVolume)

#define IWaveMixer_SetPan(This,nID,nPan)	\
    (This)->lpVtbl -> SetPan(This,nID,nPan)

#define IWaveMixer_Stop(This,nID)	\
    (This)->lpVtbl -> Stop(This,nID)

#define IWaveMixer_SetFrequency(This,nID,nFrequency)	\
    (This)->lpVtbl -> SetFrequency(This,nID,nFrequency)

#define IWaveMixer_ResetFrequency(This,nID)	\
    (This)->lpVtbl -> ResetFrequency(This,nID)

#define IWaveMixer_GetFrequency(This,nID,pnFrequency)	\
    (This)->lpVtbl -> GetFrequency(This,nID,pnFrequency)

#define IWaveMixer_GetPan(This,nID,pnPan)	\
    (This)->lpVtbl -> GetPan(This,nID,pnPan)

#define IWaveMixer_GetVolume(This,nID,pnVolume)	\
    (This)->lpVtbl -> GetVolume(This,nID,pnVolume)

#define IWaveMixer_get_UsedMemory(This,pVal)	\
    (This)->lpVtbl -> get_UsedMemory(This,pVal)

#define IWaveMixer_get_MaxMemory(This,pVal)	\
    (This)->lpVtbl -> get_MaxMemory(This,pVal)

#define IWaveMixer_put_MaxMemory(This,newVal)	\
    (This)->lpVtbl -> put_MaxMemory(This,newVal)

#define IWaveMixer_GetName(This,nID,pbstrName)	\
    (This)->lpVtbl -> GetName(This,nID,pbstrName)

#define IWaveMixer_StopAll(This)	\
    (This)->lpVtbl -> StopAll(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Add_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ BSTR bstrFileName,
    /* [retval][out] */ long __RPC_FAR *pnID);


void __RPC_STUB IWaveMixer_Add_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Clear_Proxy( 
    IWaveMixer __RPC_FAR * This);


void __RPC_STUB IWaveMixer_Clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IWaveMixer_get_LoadedWaves_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IWaveMixer_get_LoadedWaves_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_GetSoundInfo_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ int nIndex,
    /* [out] */ long __RPC_FAR *pnID,
    /* [out] */ BSTR __RPC_FAR *pbstrFileName,
    /* [out] */ long __RPC_FAR *pnOriginalFrequency,
    /* [out] */ long __RPC_FAR *pnChannels,
    /* [out] */ long __RPC_FAR *pnBits,
    /* [out] */ long __RPC_FAR *pnSize);


void __RPC_STUB IWaveMixer_GetSoundInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Initialize_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long hWnd);


void __RPC_STUB IWaveMixer_Initialize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Play_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [in] */ BOOL bLoop);


void __RPC_STUB IWaveMixer_Play_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Remove_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID);


void __RPC_STUB IWaveMixer_Remove_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_SetVolume_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [in] */ long nVolume);


void __RPC_STUB IWaveMixer_SetVolume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_SetPan_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [in] */ long nPan);


void __RPC_STUB IWaveMixer_SetPan_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_Stop_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID);


void __RPC_STUB IWaveMixer_Stop_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_SetFrequency_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [in] */ long nFrequency);


void __RPC_STUB IWaveMixer_SetFrequency_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_ResetFrequency_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID);


void __RPC_STUB IWaveMixer_ResetFrequency_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_GetFrequency_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [out] */ long __RPC_FAR *pnFrequency);


void __RPC_STUB IWaveMixer_GetFrequency_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_GetPan_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [out] */ long __RPC_FAR *pnPan);


void __RPC_STUB IWaveMixer_GetPan_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_GetVolume_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [out] */ long __RPC_FAR *pnVolume);


void __RPC_STUB IWaveMixer_GetVolume_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IWaveMixer_get_UsedMemory_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IWaveMixer_get_UsedMemory_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IWaveMixer_get_MaxMemory_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IWaveMixer_get_MaxMemory_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IWaveMixer_put_MaxMemory_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IWaveMixer_put_MaxMemory_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_GetName_Proxy( 
    IWaveMixer __RPC_FAR * This,
    /* [in] */ long nID,
    /* [out] */ BSTR __RPC_FAR *pbstrName);


void __RPC_STUB IWaveMixer_GetName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IWaveMixer_StopAll_Proxy( 
    IWaveMixer __RPC_FAR * This);


void __RPC_STUB IWaveMixer_StopAll_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IWaveMixer_INTERFACE_DEFINED__ */



#ifndef __DLSMEDIALib_LIBRARY_DEFINED__
#define __DLSMEDIALib_LIBRARY_DEFINED__

/* library DLSMEDIALib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_DLSMEDIALib;

#ifndef ___IWaveMixerEvents_DISPINTERFACE_DEFINED__
#define ___IWaveMixerEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IWaveMixerEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IWaveMixerEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("4B402B69-C8E1-4872-9BC9-2A1FF935F19E")
    _IWaveMixerEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IWaveMixerEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _IWaveMixerEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _IWaveMixerEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _IWaveMixerEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _IWaveMixerEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _IWaveMixerEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _IWaveMixerEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _IWaveMixerEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _IWaveMixerEventsVtbl;

    interface _IWaveMixerEvents
    {
        CONST_VTBL struct _IWaveMixerEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IWaveMixerEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IWaveMixerEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IWaveMixerEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IWaveMixerEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IWaveMixerEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IWaveMixerEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IWaveMixerEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IWaveMixerEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_WaveMixer;

#ifdef __cplusplus

class DECLSPEC_UUID("94326DCC-A647-497C-B2C4-69692A014513")
WaveMixer;
#endif
#endif /* __DLSMEDIALib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
