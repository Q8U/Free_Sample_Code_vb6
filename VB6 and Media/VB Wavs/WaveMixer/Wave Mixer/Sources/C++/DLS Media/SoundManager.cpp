// SoundManager.cpp: implementation of the CSoundManager class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "SoundManager.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CSoundManager::CSoundManager()
{
	m_hWnd			= NULL;
	m_nCountID		= 0;
	m_bInitialized	= false;
	m_lpDS			= NULL;
	m_nMemoryUsed   = 0;
	m_nMaxMemory	= SOUNDMANAGER_MAX_MEMORY;
}

CSoundManager::~CSoundManager()
{
	// Clear internal data
	Clear();

	// Reset internal members
	m_hWnd			= NULL;
	m_bInitialized	= false;

	// Remove refernce to direct sound object
	if( m_lpDS ) m_lpDS->Release();
	m_lpDS = NULL;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		Initialize
// Description:		Initializes the IDirectSound object and sets the 
//					cooperative level of the class.
// PreCondition:	None
// Parameters:		
//		hWnd:		[in]	A window handle that will be used to set the 
//							cooperative level of the DirectSound object.
//							If NULL is specified, the method will try get
//							the window handle of the foreground window.
// PostCondition:	None
// Return:			A boolean True if the initialisation was successful. If
//					something went wrong, a boolean False will be returned.
bool CSoundManager::Initialize( HWND hWnd )
{
	// Declare Variables
	HRESULT hr;

	// Check if direct sound is not already initialized
	if( m_bInitialized ) return false;

	// Check if the user specified a window handle
	if( hWnd == NULL )
	{
		// No handle specified, try to create one of our own
		hWnd = GetForegroundWindow();
		if( hWnd == NULL ) return false;
	}

	// Try to create a direct sound object
	hr = DirectSoundCreate( NULL, &m_lpDS, NULL ); 
	if( hr != DS_OK ) return false;

	// Try to set the cooperation level
	hr = m_lpDS->SetCooperativeLevel( hWnd, DSSCL_NORMAL );
	if( hr != DS_OK ) 
	{
		// Initialize failed
		// Release objects and return fail code
		m_lpDS->Release();
		m_bInitialized = false;
		m_hWnd = NULL;
		return false;
	}
	else
	{
		// Initialize was successful
		// Store properties in private members and return success
		m_hWnd = hWnd;
		m_bInitialized = true;
		return true;
	}
}

//////////////////////////////////////////////////////////////////////////////
// Function:		AddSound
// Description:		Adds a sound object by loading the specified file into 
//					memory.
// PreCondition:	None
// Parameters:		
//		sFileName:	[in]	A char pointer that specifies the filename of the 
//							wave sound file. 
// PostCondition:	None
// Return:			If the sound was loaded with success, the method returns
//					the ID of the sound object. If something went wrong, or
//					if the sound file could not be loaded, a value of -1 will
//					be returned.
long CSoundManager::AddSound( char* sFileName )
{
	// Declare variables
	long	nID = 0;
	CSound*	pSnd = new CSound();

	// Try to load the sound
	if( LoadWavFile( sFileName, pSnd ))
	{
		// File loaded with success
		// Create ID and add to map
		m_nCountID++;
		nID = m_nCountID;
		pSnd->nID = nID;
		m_sounds[nID] = pSnd;
	}
	else
	{
		// File load fails
		// Release memory and set error
		nID = -1;
		delete pSnd;
	}

	// Return ID
	return nID;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		LoadWavFile
// Description:		Loads the specified wave sound file into the specified 
//					sound object.
// PreCondition:	None
// Parameters:		
//		sFileName:	[in]	A char pointer that specifies the filename of the 
//							wave sound file. 
//		pSnd:		[in]	A pointer to a sound object that will receive the
//							the sound arguments if the load is successful.
// PostCondition:	None
// Return:			A boolean True if the load of the wave file was 
//					successful. If something went wrong, a boolean False will 
//					be returned.
bool CSoundManager::LoadWavFile( char* sFileName, 
								 CSoundManager::CSound* pSnd )
{
	// Declare Variables
	HRESULT				hr;
	HMMIO				hwav;  
	MMCKINFO			parent, child;
	UCHAR				*snd_buffer,
						*audio_ptr_1=NULL,
						*audio_ptr_2=NULL;
	DWORD				audio_length_1 = 0,
						audio_length_2 = 0;
	LPDIRECTSOUNDBUFFER	dsBuffer;
	DSBUFFERDESC		dsBD;           
	WAVEFORMATEX		pcmWF;

	// Check if direct sound has been initialized
	if( !m_bInitialized ) return false;

	// Check if a valid CSound object was passed
	if( pSnd == NULL ) return false;

	// Clear parent/child chunk info structures
	parent.ckid 	    = (FOURCC)0;
	parent.cksize 	    = 0;
	parent.fccType	    = (FOURCC)0;
	parent.dwDataOffset = 0;
	parent.dwFlags		= 0;
	child = parent;

	// Try to open the WAV file and obtain a valid handle
	hwav = mmioOpen( sFileName, NULL, MMIO_READ | MMIO_ALLOCBUF );
	if( hwav == NULL ) return false;

	// Descend into the RIFF 
	parent.fccType = mmioFOURCC('W', 'A', 'V', 'E');
	if( mmioDescend( hwav, &parent, NULL, MMIO_FINDRIFF ) )
    {
		// Close the file and return false -> not a WAV file!
		mmioClose( hwav, 0 );
		return false;
    }

	// Descend to the WAVEfmt 
	child.ckid = mmioFOURCC('f', 'm', 't', ' ');
	if( mmioDescend( hwav, &child, &parent, 0 ) )
    {
		// Close file and return false -> no format specified!
		mmioClose( hwav, 0 );
		return false;
    } 

	// Now read the wave format information from file
	if( mmioRead( hwav, (char *)&pcmWF, sizeof(pcmWF) ) != sizeof(pcmWF) )
    {
		// Close file and return false -> no sound data present in file!
		mmioClose( hwav, 0 );
		return false;
    }

	// Make sure that the data format is PCM
	if( pcmWF.wFormatTag != WAVE_FORMAT_PCM )
    {
		// Closer file and return false -> data is not in PCM format!
		mmioClose( hwav, 0 );
		return false;
    }

	// Now ascend up one level, so we can access data chunk
	if( mmioAscend( hwav, &child, 0 ) )
	{
		// Close file and return false -> cannot allocate data chunk!
		mmioClose( hwav, 0 );
		return false;
	}

	// Descend to the data chunk 
	child.ckid = mmioFOURCC('d', 'a', 't', 'a');
	if( mmioDescend( hwav, &child, &parent, MMIO_FINDCHUNK ) )
    {
		// Close file and return false -> cannot find data chunk!
	    mmioClose( hwav, 0 );
	    return false;
    }

	// Read the data in and set up directsound buffer
	// Allocate the memory to load sound data
	snd_buffer = (UCHAR *)malloc( child.cksize );

	// Read the wave data 
	mmioRead( hwav, (char *)snd_buffer, child.cksize);

	// Close the file
	mmioClose(hwav, 0);

	// Check if we stay within memory limit
	if( ( m_nMaxMemory > 0 ) &&
		(( m_nMemoryUsed + child.cksize ) > m_nMaxMemory ) )
	{
		// Memory limit exceeded
		free( snd_buffer );
		return false;
	}

	// Prepare to create sounds buffer
	pcmWF.cbSize		= 0;
	dsBD.dwSize			= sizeof( DSBUFFERDESC );
	dsBD.dwFlags		= DSBCAPS_CTRLDEFAULT | 
						  DSBCAPS_LOCSOFTWARE |
						  DSBCAPS_STATIC;;
	dsBD.dwBufferBytes	= child.cksize;
	dsBD.lpwfxFormat	= &pcmWF;
	dsBD.dwReserved		= 0;

	// Try to create the sound buffer
	hr = m_lpDS->CreateSoundBuffer( &dsBD, &dsBuffer, NULL );
	if( hr != DD_OK )
	{
		// Cannot allocate a sound buffer, free memory and exit
		free( snd_buffer );
		return false;
	}

	// Copy data into sound buffer
	if( FAILED( dsBuffer->Lock( 0,					 
							    child.cksize,			
								(void **) &audio_ptr_1, 
								&audio_length_1,
								(void **)&audio_ptr_2, 
								&audio_length_2,
								DSBLOCK_FROMWRITECURSOR) ) )
		// Unable to copy data into sound buffer
		return false;

	// Copy sections of circular buffers
	memcpy( audio_ptr_1, snd_buffer, audio_length_1 );
	memcpy( audio_ptr_2, (snd_buffer+audio_length_1), audio_length_2 );

	// Unlock the buffer
	if( FAILED( dsBuffer->Unlock( audio_ptr_1, 
						  		  audio_length_1, 
								  audio_ptr_2, 
								  audio_length_2 ) ) )
		// Unable to release the buffer
		return false;

	// Release the temporary buffer
	free(snd_buffer);

	// Set default settings of a sound buffer
	dsBuffer->SetPan( DSBPAN_CENTER );
	dsBuffer->SetVolume( DSBVOLUME_MAX );
	dsBuffer->SetFrequency( pcmWF.nSamplesPerSec );

	// Strip the drive, path and extension form filename
	char	sDrive[1024];
	char	sPath[1024];
	char	sFile[1024];
	char	sExtension[1024];
	_splitpath( sFileName, sDrive, sPath, sFile, sExtension );
	
	// Copy the arguments into the CSound class
	pSnd->dsBD		= dsBD;
	pSnd->dsBuffer	= dsBuffer;
	pSnd->pcmWF		= pcmWF;
	strcpy( pSnd->sFileName, sFile );
	pSnd->nSize		= dsBD.dwBufferBytes;
	pSnd->nOriginalFrequency = pcmWF.nSamplesPerSec;

	// Increment memory counter
	m_nMemoryUsed += dsBD.dwBufferBytes;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetSound
// Description:		Get a sound object with the specified ID from the sounds
//					array.
// PreCondition:	None
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
// PostCondition:	None
// Return:			A pointer to a sound object with the specified ID. If no
//					object could be found with the specified ID, a NULL
//					pointer will be returned.
CSoundManager::CSound* CSoundManager::GetSound( long nID )
{
	// Declare Variables
	CSound*	pSnd;
	MAPSOUND::iterator it;

	// Try to allocate requested object in map
	it = m_sounds.find( nID );

	// Object with specified ID found?
	if( it != m_sounds.end() )
		// Object found, return pointer
		pSnd = (*it).second;
	else
		// Object not found
		pSnd = NULL;

	// Return object (if found)
	return pSnd;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		Play
// Description:		Plays the sound with the specified ID.
// PreCondition:	None
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		bLoop:		[in]	A boolean that specifies if the sound should be
//							played in loop mode
// PostCondition:	None
// Return:			A boolean True if a sound with the specified ID could
//					be found and played. If no sound with the specified ID
//					could be found, or if an error occurs during playback, a
//					boolean False will be returned.
bool CSoundManager::Play( long nID, bool bLoop )
{
	// Declare Variables
	HRESULT hr;
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Set cursor to begin of buffer
	hr = pSnd->dsBuffer->SetCurrentPosition( 0 );
	if( hr != DS_OK ) return false;

	// Play the wav file
	if( bLoop )
		hr = pSnd->dsBuffer->Play( 0, 0, DSBPLAY_LOOPING );
	else
		hr = pSnd->dsBuffer->Play( 0, 0, 0 );
	if( hr != DS_OK ) return false;

	// Return success code
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetNumberOfSounds
// Description:		Gets the number of sound files loaded in the sounds array.
// PreCondition:	None
// Parameters:		None
// PostCondition:	None
// Return:			A long that returns the number of loaded sound files. If
//					no sound files are loaded, the method will return 0.
int CSoundManager::GetNumberOfSounds()
{
	// Return size of map object
	return m_sounds.size();
}

//////////////////////////////////////////////////////////////////////////////
// Function:		SetPan
// Description:		Sets the panning (i.e. balance) of a specified sound
//					object.
// PreCondition:	None	
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		nPan:		[in]	A long that specifies the pan value
// PostCondition:	None
// Return:			A boolean True if the pan value for the specified sound
//					object could be set. If something went wrong, a boolean
//					False will be returned.
bool CSoundManager::SetPan( long nID, long nPan )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Set sound pan property
    HRESULT hr = pSnd->dsBuffer->SetPan( nPan );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		SetVolume
// Description:		Sets the volume of a specified sound object.
// PreCondition:	None
// Parameters:	
//		nID:		[in]	A long that specifies a sound ID
//		nVolume:	[in]	A long that specifies the volume
// PostCondition:	None
// Return:			A boolean True if the volume for the specified sound
//					object could be set. If something went wrong, a boolean
//					False will be returned.
bool CSoundManager::SetVolume( long nID, long nVolume )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Set sound volume property
	HRESULT hr = pSnd->dsBuffer->SetVolume( -nVolume );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		SetFrequency
// Description:		Sets the frequency of a specified sound object.
// PreCondition:	None
// Parameters:	
//		nID:		[in]	A long that specifies a sound ID
//		nFrequency:	[in]	A long that specifies the frequency
// PostCondition:	None
// Return:			A boolean True if the frequency for the specified sound
//					object could be set. If something went wrong, a boolean
//					False will be returned.
bool CSoundManager::SetFrequency( long nID, long nFrequency )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Set sound frequency property
	HRESULT hr = pSnd->dsBuffer->SetFrequency( nFrequency );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		Stop
// Description:		Stops the playback of the specified sound object.
// PreCondition:	None
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
// PostCondition:	None
// Return:			A boolean True if the playback of the specified sound
//					object could be stopped. If something went wrong, a 
//					boolean False will be returned.
bool CSoundManager::Stop( long nID )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Stop playing of the sound
	HRESULT hr = pSnd->dsBuffer->Stop();
	if( hr != DS_OK ) return false;

	// Return success 
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		Clear
// Description:		Clears the entire contents of this class and resets the 
//					internal members.
// PreCondition:	None
// Parameters:		None
// PostCondition:	None
// Return:			None
void CSoundManager::Clear()
{
	// Declare Variables
	CSound* pSnd = NULL;

	// Clear all the sound files
	while( m_sounds.size() > 0 )
	{
		pSnd = (*m_sounds.begin()).second;
		RemoveSound( pSnd->nID );
	}

	// Reset internal members
	m_nMemoryUsed   = 0;
	m_nCountID		= 0;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		RemoveSound
// Description:		Removes the specified sound from the sounds array and 
//					releases all the used references off this object.
// PreCondition:	None
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
// PostCondition:	None
// Return:			A boolean True if the specified sound could be removed.
//					If something went wrong, a boolean False will be returned.
bool CSoundManager::RemoveSound( long nID )
{
	// Declare Variables
	CSound*	pSnd;
	MAPSOUND::iterator it;

	// Try to allocate requested object in map
	it = m_sounds.find( nID );

	// Object with specified ID found?
	if( it == m_sounds.end() ) return false;

	// Get sound object found
	pSnd = (*it).second;
	if( pSnd == NULL ) return false;

	// Remove sound object from array
	m_sounds.erase( it );
	
	// Try to release the soundbuffer
	pSnd->dsBuffer->Release();

	// Adjust memory counter
	m_nMemoryUsed -= pSnd->nSize;

	// Delete the sound object 
	delete pSnd;

	// Return success code
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetSoundInfo
// Description:		Gets detailed information about the sound at the specified
//					index in the array. 
// PreCondition:	None
// Parameters:
//		nIndex:			[in]	An integer that specifies an index in the
//								sounds array. This value must be between 1 and
//								NumberOfSounds		
//		pnID:			[out]	A long pointer that will receive the ID
//		psFileName:		[out]	A char pointer that will receive the filename
//		pnOriginalFrequency:
//						[out]	A long pointer that will recieve the original
//								(or sample) frequency from the sound
//								of the loaded wave file
//		pnChannels:		[out]	A long pointer that will receive the number
//								of used channels (i.e. 1 = Mono, 2 = Stereo)
//		pnBits:			[out]	A long pointer that will receive the number
//								bits used (e.. 8 or 16 bit)
//		pnSize:			[out]	A long pointer that will receive the size of 
//								the wave data in bytes
// PostCondition:	None
// Return:			A boolean True if the specified sound could be found and 
//					its properties returned. If something went wrong, or if
//					the index is invalid, a boolean False will be returned.
bool CSoundManager::GetSoundInfo( int    nIndex,
								  long*  pnID,
								  char*  psFileName,
								  long*  pnOriginalFrequency,
								  long*  pnChannels,
								  long*  pnBits, 
								  long*  pnSize )
{
	// Declare and initialize variables
	*pnID = 0;
	*pnSize = 0;
	*pnOriginalFrequency = 0;
	*pnChannels = 0;
	*pnBits = 0;

	// Check if the index is valid
	if( ( nIndex < 1 ) || ( nIndex > m_sounds.size() ) )
		return false;
	
	// Get the sound objectat the specified index
	MAPSOUND::iterator it = m_sounds.begin();
	advance( it, ( nIndex - 1 ) );
	CSound* pSnd = (*it).second;

	// Check if a valid sound is found
	if( pSnd == NULL ) return false;

	// Get the sound properties
	*pnID = pSnd->nID;
	strcpy( psFileName, pSnd->sFileName );
	*pnChannels = pSnd->pcmWF.nChannels;
	*pnBits = pSnd->pcmWF.wBitsPerSample;
	*pnSize = pSnd->nSize;
	*pnOriginalFrequency = pSnd->nOriginalFrequency;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetName
// Description:		Gets the file name of a specified sound.
// PreCondition:	None	
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		pnFileName:	[out]	A char pointer will the recieve the filename
// PostCondition:	None
// Return:			A boolean True if the name for the specified sound object 
//					could be returned. If something went wrong, a boolean 
//					False will be returned.
bool CSoundManager::GetName( long nID, char* psFileName )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Get sound name property
	strcpy( psFileName, pSnd->sFileName );

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		ResetFrequency
// Description:		Resets the frequency of the specified sound. This means
//					the current playback frequency will be repleced with the
//					original sample freqency of the sound.
// PreCondition:	None
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
// PostCondition:	None
// Return:			A boolean True if the specified sound frequency could be 
//					reset. If something went wrong, a boolean False will be 
//					returned.
bool CSoundManager::ResetFrequency( long nID )
{
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Get the original frequency
	long nOriginalFrequency = pSnd->nOriginalFrequency;
	if( nOriginalFrequency == 0 ) return false;

	// Set the playback frequency to the original frequency
	HRESULT hr = pSnd->dsBuffer->SetFrequency( nOriginalFrequency );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetFrequency
// Description:		Gets the current playback frequency of a specified sound.
// PreCondition:	None	
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		pnFrequency:[out]	A long will the recieve the playback frequency
// PostCondition:	None
// Return:			A boolean True if the frequency for the specified sound
//					object could be returned. If something went wrong, a 
//					boolean False will be returned.
bool CSoundManager::GetFrequency( long nID, long* pnFrequency )
{
	// Declare Variables
	*pnFrequency = 0;
	
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Get sound frequency property
	HRESULT hr = pSnd->dsBuffer->GetFrequency( (DWORD*)pnFrequency );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetPan
// Description:		Gets the current pan (i.e. balance) of a specified sound.
// PreCondition:	None	
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		pnPan:		[out]	A long will the recieve the pan
// PostCondition:	None
// Return:			A boolean True if the pan for the specified sound object 
//					could be returned. If something went wrong, a boolean 
//					False will be returned.
bool CSoundManager::GetPan( long nID, long* pnPan )
{
	// Declare Variables
	*pnPan = 0;
	
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Get sound pan property
	HRESULT hr = pSnd->dsBuffer->GetPan( pnPan );
	if( hr != DS_OK ) return false;

	// Return success
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetVolume
// Description:		Gets the current volume of a specified sound.
// PreCondition:	None	
// Parameters:		
//		nID:		[in]	A long that specifies a sound ID
//		pnVolume:	[out]	A long will the recieve the volume
// PostCondition:	None
// Return:			A boolean True if the volume for the specified sound 
//					object could be returned. If something went wrong, a 
//					boolean False will be returned.
bool CSoundManager::GetVolume( long nID, long* pnVolume )
{
	// Declare Variables
	*pnVolume = 0;
	
	// Get sound object
	CSound* pSnd = GetSound( nID );
	if( pSnd == NULL ) return false;

	// Get sound volume property
	HRESULT hr = pSnd->dsBuffer->GetVolume( pnVolume );
	if( hr != DS_OK ) return false;

	// Invert volume and return success
	*pnVolume = -(*pnVolume);
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetUsedMemory
// Description:		Gets the amount of used memory by the sound file in bytes.
//					This amount is simply the sum of all the used sound 
//					buffers by the class.
// PreCondition:	None	
// Parameters:		None
// PostCondition:	None
// Return:			A long that the specifies the amount of used memory in
//					bytes. If no buffers are allocated (i.e. no sounds 
//					loaded), this value will equal 0.
long CSoundManager::GetUsedMemory()
{
	return m_nMemoryUsed;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		GetMaxMemory
// Description:		Gets the value of the maximum allowed memory that can be
//					allocated by the soundmanager class. This value represents
//					the upper memory limit which cannot be exceeded.
// PreCondition:	None	
// Parameters:		None
// PostCondition:	None
// Return:			A long that the specifies the maximum amount of memory that
//					can be used by the current soundmanager object. This value
//					is by default set to 10Mb, however, the user can modify 
//					this limit by using the 'SetMaxMemory()' method.
long CSoundManager::GetMaxMemory()
{
	return m_nMaxMemory;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		SetMaxMemory
// Description:		Sets a new value of maximum allocatable memory by this
//					soundmanager object. This new value will first be 
//					validated: a. Must be > 0; b. Must be > current used 
//					memory; Must be < 16GB. 
// PreCondition:	None	
// Parameters:		
//		nMaxMemory:	[in]	A long that specifies the new upper limit
// PostCondition:	None
// Return:			A boolean True if the new upper memory limit could be
//					validated and stored. If something went wrong, or if 
//					the new value could not be validated, a boolean False 
//					will be returned.
bool CSoundManager::SetMaxMemory( long nMaxMemory )
{
	// Check if the new max memory specification is valid
	// (i.e. > 0 )
	if( nMaxMemory == 0 ) return 0;

	// Check if new upper limit is larger than sum of already loaded buffers
	if( nMaxMemory < m_nMemoryUsed ) return false;

	// Only allow a maximum of 16GB
	if( nMaxMemory >= 17179869184 ) return false;

	// Set the new upper memory limit
	m_nMaxMemory = nMaxMemory;
	return true;
}

//////////////////////////////////////////////////////////////////////////////
// Function:		StopAll
// Description:		Stops the playback of all loaded sound objects.
// PreCondition:	None
// Parameters:		None
// PostCondition:	None
// Return:			None
void CSoundManager::StopAll()
{
	// Declare Variables
	CSound* pSnd = NULL;
	MAPSOUND::iterator it;

	// Stop the playing of all sounds
	for( it = m_sounds.begin(); it != m_sounds.end(); it++ )
	{
		pSnd = (*it).second;
		pSnd->dsBuffer->Stop();
	}
}

