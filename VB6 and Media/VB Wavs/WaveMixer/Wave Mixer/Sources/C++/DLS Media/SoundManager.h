// SoundManager.h: interface for the CSoundManager class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SOUNDMANAGER_H__32E5B864_A1F3_4160_9537_DA9A7C4BFB91__INCLUDED_)
#define AFX_SOUNDMANAGER_H__32E5B864_A1F3_4160_9537_DA9A7C4BFB91__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// Specify the maximum amount of memory that this class can use.
// If not memory limit should be specified, set this value to 0.
#define SOUNDMANAGER_MAX_MEMORY		10485760	// 10MB Maximum

#include "mmsystem.h"
#include "dsound.h"

#include <map>
using namespace std;

class CSoundManager
{
	// Define a struct that holds all the arguments for
	// a single sound object
	class CSound
	{
		// Public properties
		public:
			long					nID;
			long					nSize;
			long					nOriginalFrequency;
			char					sFileName[1024];
			LPDIRECTSOUNDBUFFER		dsBuffer;   
			DSBUFFERDESC			dsBD;           
			WAVEFORMATEX			pcmWF;
	};

	typedef map<long, CSound*> MAPSOUND;

// Class Constructors
public:
	CSoundManager();

// Class Destructors
public:
	~CSoundManager();

// Private properties
private:
	HWND			m_hWnd;
	bool			m_bInitialized;
	LPDIRECTSOUND	m_lpDS;
	long			m_nCountID;
	MAPSOUND		m_sounds;
	long			m_nMemoryUsed;
	long			m_nMaxMemory;

// Public properties
public:

// Private methods
private:
	bool	LoadWavFile( char* sFileName, CSoundManager::CSound* pSnd );
	CSound*	GetSound( long nID );

// Public methods
public:
	bool	FromString( BSTR bstrSoundManager );			// TO DO
	void	ToString( BSTR* bstrSoundManager );				// TO DO

	long	AddSound( char* sFileName );
	long	AddSound( UINT nResourceID );
	void	Clear();
	bool	GetFrequency( long nID, long* pnFrequency );
	long	GetMaxMemory();
	bool	GetName( long nID, char* psFileName );
	int		GetNumberOfSounds();
	bool	GetPan( long nID, long* pnPan );
	bool	GetSoundInfo( int    nIndex,
						  long*  pnID,
						  char*  psFileName,
						  long*	 pnOriginalFrequency,
						  long*  pnChannels,
						  long*  pnBits,
						  long*  pnSize );
	long	GetUsedMemory();
	bool	GetVolume( long nID, long* pnVolume );
	bool	Initialize( HWND hWnd );
	bool	Play( long nID, bool bLoop );
	bool	RemoveSound( long nID );
	bool	ResetFrequency( long nID );
	bool	SetFrequency( long nID, long nFrequency );
	bool	SetMaxMemory( long nMaxMemory );
	bool	SetPan( long nID, long nPan );
	bool	SetVolume( long nID, long nVolume );
	bool	Stop( long nID );
	void	StopAll();
};

#endif // !defined(AFX_SOUNDMANAGER_H__32E5B864_A1F3_4160_9537_DA9A7C4BFB91__INCLUDED_)
