// WaveMixer.cpp : Implementation of CWaveMixer
#include "stdafx.h"
#include "DLS Media.h"
#include "WaveMixer.h"

/////////////////////////////////////////////////////////////////////////////
// CWaveMixer


STDMETHODIMP CWaveMixer::Add(BSTR bstrFileName, long* pnID)
{
	// Try to add a sound
	USES_CONVERSION;
	*pnID = m_pSndMan->AddSound( OLE2T( bstrFileName ) );
	if( *pnID == -1 )
		return E_FAIL;
	else
		return S_OK;
}

STDMETHODIMP CWaveMixer::Clear()
{
	// Clear class contents
	m_pSndMan->Clear();
	return S_OK;
}

STDMETHODIMP CWaveMixer::get_LoadedWaves(long *pVal)
{
	// Get number of loaded sounds
	*pVal = m_pSndMan->GetNumberOfSounds();
	return S_OK;
}

STDMETHODIMP CWaveMixer::GetSoundInfo( int   nIndex,
									   long* pnID, 
									   BSTR* pbstrFileName,
									   long* pnOriginalFrequency,
									   long* pnChannels,
									   long* pnBits,
									   long* pnSize)
{
	// Get information about a specified sound
	USES_CONVERSION;
	char sFileName[1024];
	bool bOk = m_pSndMan->GetSoundInfo( nIndex,
										pnID, 
										sFileName, 
										pnOriginalFrequency,
										pnChannels,
										pnBits,
										pnSize );
	
	*pbstrFileName = SysAllocString( T2OLE( sFileName ) );

	if( bOk )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::Initialize(long hWnd)
{
	// Initialize soundmanager class
	if( m_pSndMan->Initialize( (HWND)hWnd ) )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::Play(long nID, BOOL bLoop)
{
	// Declare & Initialize Variables
	bool bLooping = true;
	if( bLoop == FALSE ) bLooping = false;

	// Try to play the sound
	if( m_pSndMan->Play( nID, bLooping ) )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::Remove(long nID)
{
	// Remove the sound from the manager
	if( m_pSndMan->RemoveSound( nID ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::SetVolume(long nID, long nVolume)
{
	// Set the volume of a specified sound
	if( m_pSndMan->SetVolume( nID, nVolume ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::SetPan(long nID, long nPan)
{
	// Set the pan of a specified sound
	if( m_pSndMan->SetPan( nID, nPan ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::Stop(long nID)
{
	// Stop a specified sound
	if( m_pSndMan->Stop( nID ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::SetFrequency(long nID, long nFrequency)
{
	// Set the frequency of a specified sound
	if( m_pSndMan->SetFrequency( nID, nFrequency ) )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::ResetFrequency(long nID)
{
	// Reset the frequency to default of a specified sound
	if( m_pSndMan->ResetFrequency( nID ) )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::GetFrequency(long nID, long *pnFrequency)
{
	// Get the current playback frequency of a specified sound
	if( m_pSndMan->GetFrequency( nID, pnFrequency ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::GetPan(long nID, long *pnPan)
{
	// Get the current pan of a specified sound
	if( m_pSndMan->GetPan( nID, pnPan ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::GetVolume(long nID, long *pnVolume)
{
	// Get the current volume of a specified sound
	if( m_pSndMan->GetVolume( nID, pnVolume ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::get_UsedMemory(long *pVal)
{
	// Gets the current used memory
	*pVal = m_pSndMan->GetUsedMemory();
	return S_OK;
}

STDMETHODIMP CWaveMixer::get_MaxMemory(long *pVal)
{
	// Gets the current mamximum memory limit
	*pVal = m_pSndMan->GetMaxMemory();
	return S_OK;
}

STDMETHODIMP CWaveMixer::put_MaxMemory(long newVal)
{
	// Sets the current mamximum memory limit
	if( m_pSndMan->SetMaxMemory( newVal ))
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::GetName(long nID, BSTR *pbstrName)
{
	// Gets the name of a specific sound
	USES_CONVERSION;
	char sFileName[1024];
	bool bOk = m_pSndMan->GetName( nID,
 								   sFileName );
	*pbstrName = SysAllocString( T2OLE( sFileName ) );

	if( bOk )
		return S_OK;
	else
		return E_FAIL;
}

STDMETHODIMP CWaveMixer::StopAll()
{
	// Stops the playback of all loaded sounds
	m_pSndMan->StopAll();
	return S_OK;
}
