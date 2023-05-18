// WaveMixer.h : Declaration of the CWaveMixer

#ifndef __WAVEMIXER_H_
#define __WAVEMIXER_H_

#include "resource.h"       // main symbols
#include "SoundManager.h"

/////////////////////////////////////////////////////////////////////////////
// CWaveMixer
class ATL_NO_VTABLE CWaveMixer : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CWaveMixer, &CLSID_WaveMixer>,
	public IConnectionPointContainerImpl<CWaveMixer>,
	public IDispatchImpl<IWaveMixer, &IID_IWaveMixer, &LIBID_DLSMEDIALib>
{
public:
	CWaveMixer()
	{
		m_pSndMan = new CSoundManager();
	}

	~CWaveMixer()
	{
		if( m_pSndMan ) delete m_pSndMan;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_WAVEMIXER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CWaveMixer)
	COM_INTERFACE_ENTRY(IWaveMixer)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()
BEGIN_CONNECTION_POINT_MAP(CWaveMixer)
END_CONNECTION_POINT_MAP()


// IWaveMixer
public:
	STDMETHOD(StopAll)();
	STDMETHOD(GetName)(/*[in]*/ long nID, /*[out]*/ BSTR* pbstrName);
	STDMETHOD(get_MaxMemory)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_MaxMemory)(/*[in]*/ long newVal);
	STDMETHOD(get_UsedMemory)(/*[out, retval]*/ long *pVal);
	STDMETHOD(GetVolume)(/*[in]*/ long nID, /*[out]*/ long* pnVolume);
	STDMETHOD(GetPan)(/*[in]*/ long nID, /*[out]*/ long* pnPan);
	STDMETHOD(GetFrequency)(/*[in]*/ long nID, /*[out]*/ long* pnFrequency);
	STDMETHOD(ResetFrequency)(/*[in]*/ long nID);
	STDMETHOD(SetFrequency)(/*[in]*/ long nID, /*[in]*/ long nFrequency);
	STDMETHOD(Stop)(/*[in]*/ long nID);
	STDMETHOD(SetPan)(/*[in]*/ long nID, /*[in]*/ long nPan);
	STDMETHOD(SetVolume)(/*[in]*/ long nID, /*[in]*/ long nVolume);
	STDMETHOD(Remove)(/*[in]*/ long nID);
	STDMETHOD(Play)(/*[in]*/ long nID, /*[in]*/ BOOL bLoop);
	STDMETHOD(Initialize)(/*[in]*/ long hWnd);
	STDMETHOD(GetSoundInfo)(/*[in]*/ int nIndex, /*[out]*/ long* pnID, /*[out]*/ BSTR* pbstrFileName, /*[out]*/ long* pnOriginalFrequency, /*[out]*/ long* pnChannels, /*[out]*/ long* pnBits, /*[out]*/ long* pnSize);
	STDMETHOD(get_LoadedWaves)(/*[out, retval]*/ long *pVal);
	STDMETHOD(Clear)();
	STDMETHOD(Add)(/*[in]*/ BSTR bstrFileName, /*[out]*/ long* pnID);

// Private Member Variables
private:
	CSoundManager*	m_pSndMan;
};

#endif //__WAVEMIXER_H_
