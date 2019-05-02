Attribute VB_Name = "modBass21"
'BASS 2.1 Visual Basic API, (c) 1999-2004 Ian Luck.
'Please report bugs/suggestions/etc... to bass@un4seen.com

'See the BASS.CHM file for detailed documentation

'NOTE: VB does not support 64-bit integers, so VB users only have access
'      to the low 32-bits of 64-bit return values. 64-bit parameters can
'      be specified though, using the "64" version of the function.

'NOTE: Use the VBStrFromAnsiPtr function to convert "char *" to VB "String".


Global Const BASSTRUE As Long = 1   'Use this instead of VB Booleans
Global Const BASSFALSE As Long = 0  'Use this instead of VB Booleans

'***********************************************
'* Error codes returned by BASS_ErrorGetCode() *
'***********************************************
Global Const BASS_OK = 0               'all is OK
Global Const BASS_ERROR_MEM = 1        'memory error
Global Const BASS_ERROR_FILEOPEN = 2   'can't open the file
Global Const BASS_ERROR_DRIVER = 3     'can't find a free sound driver
Global Const BASS_ERROR_BUFLOST = 4    'the sample buffer was lost
Global Const BASS_ERROR_HANDLE = 5     'invalid handle
Global Const BASS_ERROR_FORMAT = 6     'unsupported sample format
Global Const BASS_ERROR_POSITION = 7   'invalid playback position
Global Const BASS_ERROR_INIT = 8       'BASS_Init has not been successfully called
Global Const BASS_ERROR_START = 9      'BASS_Start has not been successfully called
Global Const BASS_ERROR_ALREADY = 14   'already initialized/paused/whatever
Global Const BASS_ERROR_NOPAUSE = 16   'not paused
Global Const BASS_ERROR_NOCHAN = 18    'can't get a free channel
Global Const BASS_ERROR_ILLTYPE = 19   'an illegal type was specified
Global Const BASS_ERROR_ILLPARAM = 20  'an illegal parameter was specified
Global Const BASS_ERROR_NO3D = 21      'no 3D support
Global Const BASS_ERROR_NOEAX = 22     'no EAX support
Global Const BASS_ERROR_DEVICE = 23    'illegal device number
Global Const BASS_ERROR_NOPLAY = 24    'not playing
Global Const BASS_ERROR_FREQ = 25      'illegal sample rate
Global Const BASS_ERROR_NOTFILE = 27   'the stream is not a file stream
Global Const BASS_ERROR_NOHW = 29      'no hardware voices available
Global Const BASS_ERROR_EMPTY = 31     'the MOD music has no sequence data
Global Const BASS_ERROR_NONET = 32     'no internet connection could be opened
Global Const BASS_ERROR_CREATE = 33    'couldn't create the file
Global Const BASS_ERROR_NOFX = 34      'effects are not available
Global Const BASS_ERROR_PLAYING = 35   'the channel is playing
Global Const BASS_ERROR_NOTAVAIL = 37  'requested data is not available
Global Const BASS_ERROR_DECODE = 38    'the channel is a "decoding channel"
Global Const BASS_ERROR_DX = 39        'a sufficient DirectX version is not installed
Global Const BASS_ERROR_TIMEOUT = 40   'connection timedout
Global Const BASS_ERROR_FILEFORM = 41  'unsupported file format
Global Const BASS_ERROR_SPEAKER = 42   'unavailable speaker
Global Const BASS_ERROR_UNKNOWN = -1   'some other mystery error

'************************
'* Initialization flags *
'************************
Global Const BASS_DEVICE_8BITS = 1     'use 8 bit resolution, else 16 bit
Global Const BASS_DEVICE_MONO = 2      'use mono, else stereo
Global Const BASS_DEVICE_3D = 4        'enable 3D functionality
' If the BASS_DEVICE_3D flag is not specified when initilizing BASS,
' then the 3D flags (BASS_SAMPLE_3D and BASS_MUSIC_3D) are ignored when
' loading/creating a sample/stream/music.
Global Const BASS_DEVICE_LATENCY = 256 'calculate device latency (BASS_INFO struct)
Global Const BASS_DEVICE_SPEAKERS = 2048 'force enabling of speaker assignment

'***********************************
'* BASS_INFO flags (from DSOUND.H) *
'***********************************
Global Const DSCAPS_CONTINUOUSRATE = 16
' supports all sample rates between min/maxrate
Global Const DSCAPS_EMULDRIVER = 32
' device does NOT have hardware DirectSound support
Global Const DSCAPS_CERTIFIED = 64
' device driver has been certified by Microsoft
' The following flags tell what type of samples are supported by HARDWARE
' mixing, all these formats are supported by SOFTWARE mixing
Global Const DSCAPS_SECONDARYMONO = 256    ' mono
Global Const DSCAPS_SECONDARYSTEREO = 512  ' stereo
Global Const DSCAPS_SECONDARY8BIT = 1024   ' 8 bit
Global Const DSCAPS_SECONDARY16BIT = 2048  ' 16 bit

'*****************************************
'* BASS_RECORDINFO flags (from DSOUND.H) *
'*****************************************
Global Const DSCCAPS_EMULDRIVER = DSCAPS_EMULDRIVER
' device does NOT have hardware DirectSound recording support
Global Const DSCCAPS_CERTIFIED = DSCAPS_CERTIFIED
' device driver has been certified by Microsoft

'******************************************************************
'* defines for formats field of BASS_RECORDINFO (from MMSYSTEM.H) *
'******************************************************************
Global Const WAVE_FORMAT_1M08 = &H1          ' 11.025 kHz, Mono,   8-bit
Global Const WAVE_FORMAT_1S08 = &H2          ' 11.025 kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_1M16 = &H4          ' 11.025 kHz, Mono,   16-bit
Global Const WAVE_FORMAT_1S16 = &H8          ' 11.025 kHz, Stereo, 16-bit
Global Const WAVE_FORMAT_2M08 = &H10         ' 22.05  kHz, Mono,   8-bit
Global Const WAVE_FORMAT_2S08 = &H20         ' 22.05  kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_2M16 = &H40         ' 22.05  kHz, Mono,   16-bit
Global Const WAVE_FORMAT_2S16 = &H80         ' 22.05  kHz, Stereo, 16-bit
Global Const WAVE_FORMAT_4M08 = &H100        ' 44.1   kHz, Mono,   8-bit
Global Const WAVE_FORMAT_4S08 = &H200        ' 44.1   kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_4M16 = &H400        ' 44.1   kHz, Mono,   16-bit
Global Const WAVE_FORMAT_4S16 = &H800        ' 44.1   kHz, Stereo, 16-bit

'*********************
'* Sample info flags *
'*********************
Global Const BASS_SAMPLE_8BITS = 1          ' 8 bit
Global Const BASS_SAMPLE_FLOAT = 256        ' 32-bit floating-point
Global Const BASS_SAMPLE_MONO = 2           ' mono, else stereo
Global Const BASS_SAMPLE_LOOP = 4           ' looped
Global Const BASS_SAMPLE_3D = 8             ' 3D functionality enabled
Global Const BASS_SAMPLE_SOFTWARE = 16      ' it's NOT using hardware mixing
Global Const BASS_SAMPLE_MUTEMAX = 32       ' muted at max distance (3D only)
Global Const BASS_SAMPLE_VAM = 64           ' uses the DX7 voice allocation & management
Global Const BASS_SAMPLE_FX = 128           ' old implementation of DX8 effects are enabled
Global Const BASS_SAMPLE_OVER_VOL = &H10000 ' override lowest volume
Global Const BASS_SAMPLE_OVER_POS = &H20000 ' override longest playing
Global Const BASS_SAMPLE_OVER_DIST = &H30000 ' override furthest from listener (3D only)

Global Const BASS_MP3_SETPOS = &H20000      ' enable pin-point seeking on the MP3/MP2/MP1

Global Const BASS_STREAM_AUTOFREE = &H40000 ' automatically free the stream when it stop/ends
Global Const BASS_STREAM_RESTRATE = &H80000 ' restrict the download rate of internet file streams
Global Const BASS_STREAM_BLOCK = &H100000   ' download/play internet file stream in small blocks
Global Const BASS_STREAM_DECODE = &H200000  ' don't play the stream, only decode (BASS_ChannelGetData)
Global Const BASS_STREAM_META = &H400000    ' request metadata from a Shoutcast stream
Global Const BASS_STREAM_STATUS = &H800000  ' give server status info (HTTP/ICY tags) in DOWNLOADPROC

Global Const BASS_MUSIC_FLOAT = BASS_SAMPLE_FLOAT ' 32-bit floating-point
Global Const BASS_MUSIC_MONO = BASS_SAMPLE_MONO ' force mono mixing (less CPU usage)
Global Const BASS_MUSIC_LOOP = BASS_SAMPLE_LOOP ' loop music
Global Const BASS_MUSIC_3D = BASS_SAMPLE_3D ' enable 3D functionality
Global Const BASS_MUSIC_FX = BASS_SAMPLE_FX ' enable old implementation of DX8 effects
Global Const BASS_MUSIC_AUTOFREE = BASS_STREAM_AUTOFREE ' automatically free the music when it stop/ends
Global Const BASS_MUSIC_DECODE = BASS_STREAM_DECODE ' don't play the music, only decode (BASS_ChannelGetData)
Global Const BASS_MUSIC_RAMP = &H200        ' normal ramping
Global Const BASS_MUSIC_RAMPS = &H400       ' sensitive ramping
Global Const BASS_MUSIC_SURROUND = &H800    ' surround sound
Global Const BASS_MUSIC_SURROUND2 = &H1000  ' surround sound (mode 2)
Global Const BASS_MUSIC_FT2MOD = &H2000     ' play .MOD as FastTracker 2 does
Global Const BASS_MUSIC_PT1MOD = &H4000     ' play .MOD as ProTracker 1 does
Global Const BASS_MUSIC_CALCLEN = 32768    ' calculate playback length
Global Const BASS_MUSIC_NONINTER = &H10000  ' non-interpolated mixing
Global Const BASS_MUSIC_POSRESET = &H20000  ' stop all notes when moving position
Global Const BASS_MUSIC_POSRESETEX = &H400000 ' stop all notes and reset bmp/etc when moving position
Global Const BASS_MUSIC_STOPBACK = &H80000  ' stop the music on a backwards jump effect
Global Const BASS_MUSIC_NOSAMPLE = &H100000 ' don't load the samples

' Speaker assignment flags
Global Const BASS_SPEAKER_FRONT = &H1000000 ' front speakers
Global Const BASS_SPEAKER_REAR = &H2000000  ' rear/side speakers
Global Const BASS_SPEAKER_CENLFE = &H3000000 ' center & LFE speakers (5.1)
Global Const BASS_SPEAKER_REAR2 = &H4000000 ' rear center speakers (7.1)
Global Const BASS_SPEAKER_LEFT = &H10000000 ' modifier: left
Global Const BASS_SPEAKER_RIGHT = &H20000000 ' modifier: right
Global Const BASS_SPEAKER_FRONTLEFT = BASS_SPEAKER_FRONT Or BASS_SPEAKER_LEFT
Global Const BASS_SPEAKER_FRONTRIGHT = BASS_SPEAKER_FRONT Or BASS_SPEAKER_RIGHT
Global Const BASS_SPEAKER_REARLEFT = BASS_SPEAKER_REAR Or BASS_SPEAKER_LEFT
Global Const BASS_SPEAKER_REARRIGHT = BASS_SPEAKER_REAR Or BASS_SPEAKER_RIGHT
Global Const BASS_SPEAKER_CENTER = BASS_SPEAKER_CENLFE Or BASS_SPEAKER_LEFT
Global Const BASS_SPEAKER_LFE = BASS_SPEAKER_CENLFE Or BASS_SPEAKER_RIGHT
Global Const BASS_SPEAKER_REAR2LEFT = BASS_SPEAKER_REAR2 Or BASS_SPEAKER_LEFT
Global Const BASS_SPEAKER_REAR2RIGHT = BASS_SPEAKER_REAR2 Or BASS_SPEAKER_RIGHT

Global Const BASS_UNICODE = &H80000000

Global Const BASS_RECORD_PAUSE = 32768 ' start recording paused

'**********************************************
'* BASS_StreamGetTags flags : what's returned *
'**********************************************
Global Const BASS_TAG_ID3 = 0   'ID3v1 tags : 128 byte block
Global Const BASS_TAG_ID3V2 = 1 'ID3v2 tags : variable length block
Global Const BASS_TAG_OGG = 2   'OGG comments : array of null-terminated strings
Global Const BASS_TAG_HTTP = 3  'HTTP headers : array of null-terminated strings
Global Const BASS_TAG_ICY = 4   'ICY headers : array of null-terminated strings
Global Const BASS_TAG_META = 5  'ICY metadata : null-terminated string

'********************
'* 3D channel modes *
'********************
Global Const BASS_3DMODE_NORMAL = 0
' normal 3D processing
Global Const BASS_3DMODE_RELATIVE = 1
' The channel's 3D position (position/velocity/orientation) are relative to
' the listener. When the listener's position/velocity/orientation is changed
' with BASS_Set3DPosition, the channel's position relative to the listener does
' not change.
Global Const BASS_3DMODE_OFF = 2
' Turn off 3D processing on the channel, the sound will be played
' in the center.

'****************************************************
'* EAX environments, use with BASS_SetEAXParameters *
'****************************************************
Global Const EAX_ENVIRONMENT_GENERIC = 0
Global Const EAX_ENVIRONMENT_PADDEDCELL = 1
Global Const EAX_ENVIRONMENT_ROOM = 2
Global Const EAX_ENVIRONMENT_BATHROOM = 3
Global Const EAX_ENVIRONMENT_LIVINGROOM = 4
Global Const EAX_ENVIRONMENT_STONEROOM = 5
Global Const EAX_ENVIRONMENT_AUDITORIUM = 6
Global Const EAX_ENVIRONMENT_CONCERTHALL = 7
Global Const EAX_ENVIRONMENT_CAVE = 8
Global Const EAX_ENVIRONMENT_ARENA = 9
Global Const EAX_ENVIRONMENT_HANGAR = 10
Global Const EAX_ENVIRONMENT_CARPETEDHALLWAY = 11
Global Const EAX_ENVIRONMENT_HALLWAY = 12
Global Const EAX_ENVIRONMENT_STONECORRIDOR = 13
Global Const EAX_ENVIRONMENT_ALLEY = 14
Global Const EAX_ENVIRONMENT_FOREST = 15
Global Const EAX_ENVIRONMENT_CITY = 16
Global Const EAX_ENVIRONMENT_MOUNTAINS = 17
Global Const EAX_ENVIRONMENT_QUARRY = 18
Global Const EAX_ENVIRONMENT_PLAIN = 19
Global Const EAX_ENVIRONMENT_PARKINGLOT = 20
Global Const EAX_ENVIRONMENT_SEWERPIPE = 21
Global Const EAX_ENVIRONMENT_UNDERWATER = 22
Global Const EAX_ENVIRONMENT_DRUGGED = 23
Global Const EAX_ENVIRONMENT_DIZZY = 24
Global Const EAX_ENVIRONMENT_PSYCHOTIC = 25
' total number of environments
Global Const EAX_ENVIRONMENT_COUNT = 26

'**********************************************************************
'* Sync types (with BASS_ChannelSetSync() "param" and SYNCPROC "data" *
'* definitions) & flags.                                              *
'**********************************************************************
' Sync when a music or stream reaches a position.
' if HMUSIC...
' param: LOWORD=order (0=first, -1=all) HIWORD=row (0=first, -1=all)
' data : LOWORD=order HIWORD=row
' if HSTREAM...
' param: position in bytes
' data : not used
Global Const BASS_SYNC_POS = 0
Global Const BASS_SYNC_MUSICPOS = 0
' Sync when an instrument (sample for the non-instrument based formats)
' is played in a music (not including retrigs).
' param: LOWORD=instrument (1=first) HIWORD=note (0=c0...119=b9, -1=all)
' data : LOWORD=note HIWORD=volume (0-64)
Global Const BASS_SYNC_MUSICINST = 1
' Sync when a music or file stream reaches the end.
' param: not used
' data : not used
Global Const BASS_SYNC_END = 2
' Sync when the "sync" effect (XM/MTM/MOD: E8x/Wxx, IT/S3M: S2x) is used.
' param: 0:data=pos, 1:data="x" value
' data : param=0: LOWORD=order HIWORD=row, param=1: "x" value
Global Const BASS_SYNC_MUSICFX = 3
' FLAG: post a Windows message (instead of callback)
' When using a window message "callback", the message to post is given in the "proc"
' parameter of BASS_ChannelSetSync, and is posted to the window specified in the BASS_Init
' call. The message parameters are: WPARAM = data, LPARAM = user.
Global Const BASS_SYNC_META = 4
' Sync when metadata is received in a Shoutcast stream.
' param: not used
' data : pointer to the metadata
Global Const BASS_SYNC_SLIDE = 5
' Sync when an attribute slide is completed.
' param: not used
' data : the type of slide completed (one of the BASS_SLIDE_xxx values)
Global Const BASS_SYNC_STALL = 6
' Sync when playback has stalled.
' param: not used
' data : 0=stalled, 1=resumed
Global Const BASS_SYNC_DOWNLOAD = 7
' Sync when downloading of an internet (or "buffered" user file) stream has ended.
' param: not used
' data : not used
Global Const BASS_SYNC_MESSAGE = &H20000000
'FLAG: sync at mixtime, else at playtime
Global Const BASS_SYNC_MIXTIME = &H40000000
' FLAG: sync only once, else continuously
Global Const BASS_SYNC_ONETIME = &H80000000

' BASS_ChannelIsActive return values
Global Const BASS_ACTIVE_STOPPED = 0
Global Const BASS_ACTIVE_PLAYING = 1
Global Const BASS_ACTIVE_STALLED = 2
Global Const BASS_ACTIVE_PAUSED = 3

' BASS_ChannelIsSliding return flags
Global Const BASS_SLIDE_FREQ = 1
Global Const BASS_SLIDE_VOL = 2
Global Const BASS_SLIDE_PAN = 4

' BASS_ChannelGetData flags
Global Const BASS_DATA_AVAILABLE = 0         ' query how much data is buffered
Global Const BASS_DATA_FFT512 = &H80000000   ' 512 sample FFT
Global Const BASS_DATA_FFT1024 = &H80000001  ' 1024 FFT
Global Const BASS_DATA_FFT2048 = &H80000002  ' 2048 FFT
Global Const BASS_DATA_FFT4096 = &H80000003  ' 4096 FFT
Global Const BASS_DATA_FFT_INDIVIDUAL = &H10 ' FFT flag: FFT for each channel, else all combined
Global Const BASS_DATA_FFT_NOWINDOW = &H20   ' FFT flag: no Hanning window

' BASS_RecordSetInput flags
Global Const BASS_INPUT_OFF = &H10000
Global Const BASS_INPUT_ON = &H20000
Global Const BASS_INPUT_LEVEL = &H40000

Global Const BASS_INPUT_TYPE_MASK = &HFF000000
Global Const BASS_INPUT_TYPE_UNDEF = &H0
Global Const BASS_INPUT_TYPE_DIGITAL = &H1000000
Global Const BASS_INPUT_TYPE_LINE = &H2000000
Global Const BASS_INPUT_TYPE_MIC = &H3000000
Global Const BASS_INPUT_TYPE_SYNTH = &H4000000
Global Const BASS_INPUT_TYPE_CD = &H5000000
Global Const BASS_INPUT_TYPE_PHONE = &H6000000
Global Const BASS_INPUT_TYPE_SPEAKER = &H7000000
Global Const BASS_INPUT_TYPE_WAVE = &H8000000
Global Const BASS_INPUT_TYPE_AUX = &H9000000
Global Const BASS_INPUT_TYPE_ANALOG = &HA000000

' BASS_MusicSet/GetAttribute options
Global Const BASS_MUSIC_ATTRIB_AMPLIFY = 0
Global Const BASS_MUSIC_ATTRIB_PANSEP = 1
Global Const BASS_MUSIC_ATTRIB_PSCALER = 2
Global Const BASS_MUSIC_ATTRIB_BPM = 3
Global Const BASS_MUSIC_ATTRIB_SPEED = 4
Global Const BASS_MUSIC_ATTRIB_VOL_GLOBAL = 5
Global Const BASS_MUSIC_ATTRIB_VOL_CHAN = &H100 ' + channel #
Global Const BASS_MUSIC_ATTRIB_VOL_INST = &H200 ' + instrument #

' BASS_Set/GetConfig options
Global Const BASS_CONFIG_BUFFER = 0
Global Const BASS_CONFIG_UPDATEPERIOD = 1
Global Const BASS_CONFIG_MAXVOL = 3
Global Const BASS_CONFIG_GVOL_SAMPLE = 4
Global Const BASS_CONFIG_GVOL_STREAM = 5
Global Const BASS_CONFIG_GVOL_MUSIC = 6
Global Const BASS_CONFIG_CURVE_VOL = 7
Global Const BASS_CONFIG_CURVE_PAN = 8
Global Const BASS_CONFIG_FLOATDSP = 9
Global Const BASS_CONFIG_3DALGORITHM = 10
Global Const BASS_CONFIG_NET_TIMEOUT = 11
Global Const BASS_CONFIG_NET_BUFFER = 12
Global Const BASS_CONFIG_PAUSE_NOPLAY = 13
Global Const BASS_CONFIG_NET_NOPROXY = 14

' BASS_StreamGetFilePosition modes
Global Const BASS_FILEPOS_DECODE = 0
Global Const BASS_FILEPOS_DOWNLOAD = 1
Global Const BASS_FILEPOS_END = 2

' STREAMFILEPROC actions
Global Const BASS_FILE_CLOSE = 0
Global Const BASS_FILE_READ = 1
Global Const BASS_FILE_QUERY = 2
Global Const BASS_FILE_LEN = 3

Global Const BASS_STREAMPROC_END = &H80000000 ' end of user stream flag

'**************************************************************
'* DirectSound interfaces (for use with BASS_GetDSoundObject) *
'**************************************************************
Global Const BASS_OBJECT_DS = 1                     ' DirectSound
Global Const BASS_OBJECT_DS3DL = 2                  'IDirectSound3DListener

'******************************
'* DX7 voice allocation flags *
'******************************
' Play the sample in hardware. If no hardware voices are available then
' the "play" call will fail
Global Const BASS_VAM_HARDWARE = 1
' Play the sample in software (ie. non-accelerated). No other VAM flags
'may be used together with this flag.
Global Const BASS_VAM_SOFTWARE = 2

'******************************
'* DX7 voice management flags *
'******************************
' These flags enable hardware resource stealing... if the hardware has no
' available voices, a currently playing buffer will be stopped to make room for
' the new buffer. NOTE: only samples loaded/created with the BASS_SAMPLE_VAM
' flag are considered for termination by the DX7 voice management.

' If there are no free hardware voices, the buffer to be terminated will be
' the one with the least time left to play.
Global Const BASS_VAM_TERM_TIME = 4
' If there are no free hardware voices, the buffer to be terminated will be
' one that was loaded/created with the BASS_SAMPLE_MUTEMAX flag and is beyond
' it 's max distance. If there are no buffers that match this criteria, then the
' "play" call will fail.
Global Const BASS_VAM_TERM_DIST = 8
' If there are no free hardware voices, the buffer to be terminated will be
' the one with the lowest priority.
Global Const BASS_VAM_TERM_PRIO = 16

'**********************************************************************
'* software 3D mixing algorithm modes (used with BASS_Set3DAlgorithm) *
'**********************************************************************
' default algorithm (currently translates to BASS_3DALG_OFF)
Global Const BASS_3DALG_DEFAULT = 0
' Uses normal left and right panning. The vertical axis is ignored except for
'scaling of volume due to distance. Doppler shift and volume scaling are still
'applied, but the 3D filtering is not performed. This is the most CPU efficient
'software implementation, but provides no virtual 3D audio effect. Head Related
'Transfer Function processing will not be done. Since only normal stereo panning
'is used, a channel using this algorithm may be accelerated by a 2D hardware
'voice if no free 3D hardware voices are available.
Global Const BASS_3DALG_OFF = 1
' This algorithm gives the highest quality 3D audio effect, but uses more CPU.
' Requires Windows 98 2nd Edition or Windows 2000 that uses WDM drivers, if this
' mode is not available then BASS_3DALG_OFF will be used instead.
Global Const BASS_3DALG_FULL = 2
' This algorithm gives a good 3D audio effect, and uses less CPU than the FULL
' mode. Requires Windows 98 2nd Edition or Windows 2000 that uses WDM drivers, if
' this mode is not available then BASS_3DALG_OFF will be used instead.
Global Const BASS_3DALG_LIGHT = 3

Type BASS_INFO
    size As Long          ' size of this struct (set this before calling the function)
    flags As Long         ' device capabilities (DSCAPS_xxx flags)
    hwsize As Long        ' size of total device hardware memory
    hwfree As Long        ' size of free device hardware memory
    freesam As Long       ' number of free sample slots in the hardware
    free3d As Long        ' number of free 3D sample slots in the hardware
    minrate As Long       ' min sample rate supported by the hardware
    maxrate As Long       ' max sample rate supported by the hardware
    eax As Long           ' device supports EAX? (always BASSFALSE if BASS_DEVICE_3D was not used)
    minbuf As Long        ' recommended minimum buffer length in ms (requires BASS_DEVICE_LATENCY)
    dsver As Long         ' DirectSound version
    latency As Long       ' delay (in ms) before start of playback (requires BASS_DEVICE_LATENCY)
    initflags As Long     ' "flags" parameter of BASS_Init call
    speakers As Long      ' number of speakers available
    driver As Long        ' driver
End Type

Type BASS_RECORDINFO
    size As Long          ' size of this struct (set this before calling the function)
    flags As Long         ' device capabilities (DSCCAPS_xxx flags)
    formats As Long       ' supported standard formats (WAVE_FORMAT_xxx flags)
    inputs As Long        ' number of inputs
    singlein As Long      ' BASSTRUE = only 1 input can be set at a time
    driver As Long        ' driver
End Type

Type BASS_SAMPLE
    freq As Long          ' default playback rate
    volume As Long        ' default volume (0-100)
    pan As Long           ' default pan (-100=left, 0=middle, 100=right)
    flags As Long         ' BASS_SAMPLE_xxx flags
    length As Long        ' length (in samples, not bytes)
    max As Long           ' maximum simultaneous playbacks
        origres As Long       ' original resolution
    ' The following are the sample's default 3D attributes (if the sample
    ' is 3D, BASS_SAMPLE_3D is in flags) see BASS_ChannelSet3DAttributes
    mode3d As Long        ' BASS_3DMODE_xxx mode
    mindist As Single     ' minimum distance
    MAXDIST As Single     ' maximum distance
    iangle As Long        ' angle of inside projection cone
    oangle As Long        ' angle of outside projection cone
    outvol As Long        ' delta-volume outside the projection cone
    ' The following are the defaults used if the sample uses the DirectX 7
    ' voice allocation/management features.
    vam As Long           ' voice allocation/management flags (BASS_VAM_xxx)
    priority As Long      ' priority (0=lowest, &Hffffffff=highest)
End Type

Type BASS_CHANNELINFO
        freq As Long          ' default playback rate
        chans As Long         ' channels
        flags As Long         ' BASS_SAMPLE/STREAM/MUSIC/SPEAKER flags
        ctype As Long         ' type of channel
        origres As Long       ' original resolution
End Type

' BASS_CHANNELINFO types
Global Const BASS_CTYPE_SAMPLE = 1
Global Const BASS_CTYPE_RECORD = 2
Global Const BASS_CTYPE_STREAM = &H10000
Global Const BASS_CTYPE_STREAM_WAV = &H10001
Global Const BASS_CTYPE_STREAM_OGG = &H10002
Global Const BASS_CTYPE_STREAM_MP1 = &H10003
Global Const BASS_CTYPE_STREAM_MP2 = &H10004
Global Const BASS_CTYPE_STREAM_MP3 = &H10005
Global Const BASS_CTYPE_MUSIC_MOD = &H20000
Global Const BASS_CTYPE_MUSIC_MTM = &H20001
Global Const BASS_CTYPE_MUSIC_S3M = &H20002
Global Const BASS_CTYPE_MUSIC_XM = &H20003
Global Const BASS_CTYPE_MUSIC_IT = &H20004
Global Const BASS_CTYPE_MUSIC_MO3 = &H100    ' mo3 flag

'********************************************************
'* 3D vector (for 3D positions/velocities/orientations) *
'********************************************************
Type BASS_3DVECTOR
    X As Single           ' +=right, -=left
    Y As Single           ' +=up, -=down
    z As Single           ' +=front, -=behind
End Type

' DX8 effect types, use with BASS_ChannelSetFX
Global Const BASS_FX_CHORUS = 0         ' GUID_DSFX_STANDARD_CHORUS
Global Const BASS_FX_COMPRESSOR = 1     ' GUID_DSFX_STANDARD_COMPRESSOR
Global Const BASS_FX_DISTORTION = 2     ' GUID_DSFX_STANDARD_DISTORTION
Global Const BASS_FX_ECHO = 3           ' GUID_DSFX_STANDARD_ECHO
Global Const BASS_FX_FLANGER = 4        ' GUID_DSFX_STANDARD_FLANGER
Global Const BASS_FX_GARGLE = 5         ' GUID_DSFX_STANDARD_GARGLE
Global Const BASS_FX_I3DL2REVERB = 6    ' GUID_DSFX_STANDARD_I3DL2REVERB
Global Const BASS_FX_PARAMEQ = 7        ' GUID_DSFX_STANDARD_PARAMEQ
Global Const BASS_FX_REVERB = 8         ' GUID_DSFX_WAVES_REVERB

Type BASS_FXCHORUS              ' DSFXChorus
    fWetDryMix As Single
    fDepth As Single
    fFeedback As Single
    fFrequency As Single
    lWaveform As Long   ' 0=triangle, 1=sine
    fDelay As Single
    lPhase As Long              ' BASS_FX_PHASE_xxx
End Type

Type BASS_FXCOMPRESSOR  ' DSFXCompressor
    fGain As Single
    fAttack As Single
    fRelease As Single
    fThreshold As Single
    fRatio As Single
    fPredelay As Single
End Type

Type BASS_FXDISTORTION  ' DSFXDistortion
    fGain As Single
    fEdge As Single
    fPostEQCenterFrequency As Single
    fPostEQBandwidth As Single
    fPreLowpassCutoff As Single
End Type

Type BASS_FXECHO                ' DSFXEcho
    fWetDryMix As Single
    fFeedback As Single
    fLeftDelay As Single
    fRightDelay As Single
    lPanDelay As Long
End Type

Type BASS_FXFLANGER             ' DSFXFlanger
    fWetDryMix As Single
    fDepth As Single
    fFeedback As Single
    fFrequency As Single
    lWaveform As Long   ' 0=triangle, 1=sine
    fDelay As Single
    lPhase As Long              ' BASS_FX_PHASE_xxx
End Type

Type BASS_FXGARGLE              ' DSFXGargle
    dwRateHz As Long               ' Rate of modulation in hz
    dwWaveShape As Long            ' 0=triangle, 1=square
End Type

Type BASS_FXI3DL2REVERB ' DSFXI3DL2Reverb
    lRoom As Long                    ' [-10000, 0]      default: -1000 mB
    lRoomHF As Long                  ' [-10000, 0]      default: 0 mB
    flRoomRolloffFactor As Single    ' [0.0, 10.0]      default: 0.0
    flDecayTime As Single            ' [0.1, 20.0]      default: 1.49s
    flDecayHFRatio As Single         ' [0.1, 2.0]       default: 0.83
    lReflections As Long             ' [-10000, 1000]   default: -2602 mB
    flReflectionsDelay As Single     ' [0.0, 0.3]       default: 0.007 s
    lReverb As Long                  ' [-10000, 2000]   default: 200 mB
    flReverbDelay As Single          ' [0.0, 0.1]       default: 0.011 s
    flDiffusion As Single            ' [0.0, 100.0]     default: 100.0 %
    flDensity As Single              ' [0.0, 100.0]     default: 100.0 %
    flHFReference As Single          ' [20.0, 20000.0]  default: 5000.0 Hz
End Type

Type BASS_FXPARAMEQ             ' DSFXParamEq
    fCenter As Single
    fBandwidth As Single
    fGain As Single
End Type

Type BASS_FXREVERB              ' DSFXWavesReverb
    fInGain As Single                ' [-96.0,0.0]            default: 0.0 dB
    fReverbMix As Single             ' [-96.0,0.0]            default: 0.0 db
    fReverbTime As Single            ' [0.001,3000.0]         default: 1000.0 ms
    fHighFreqRTRatio As Single       ' [0.001,0.999]          default: 0.001
End Type

Global Const BASS_FX_PHASE_NEG_180 = 0
Global Const BASS_FX_PHASE_NEG_90 = 1
Global Const BASS_FX_PHASE_ZERO = 2
Global Const BASS_FX_PHASE_90 = 3
Global Const BASS_FX_PHASE_180 = 4

Type GUID       ' used with BASS_Init - use VarPtr(guid) in clsid parameter
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Declare Function BASS_SetConfig Lib "bass21.dll" (ByVal opt As Long, ByVal value As Long) As Long
Declare Function BASS_GetConfig Lib "bass21.dll" (ByVal opt As Long) As Long
Declare Function BASS_GetVersion Lib "bass21.dll" () As Long
Declare Function BASS_GetDeviceDescription Lib "bass21.dll" (ByVal device As Long) As Long
Declare Function BASS_ErrorGetCode Lib "bass21.dll" () As Long
Declare Function BASS_Init Lib "bass21.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long, ByVal win As Long, ByVal clsid As Long) As Long
Declare Function BASS_SetDevice Lib "bass21.dll" (ByVal device As Long) As Long
Declare Function BASS_GetDevice Lib "bass21.dll" () As Long
Declare Function BASS_Free Lib "bass21.dll" () As Long
Declare Function BASS_GetDSoundObject Lib "bass21.dll" (ByVal object As Long) As Long
Declare Function BASS_GetInfo Lib "bass21.dll" (ByRef info As BASS_INFO) As Long
Declare Function BASS_Update Lib "bass21.dll" () As Long
Declare Function BASS_GetCPU Lib "bass21.dll" () As Single
Declare Function BASS_Start Lib "bass21.dll" () As Long
Declare Function BASS_Stop Lib "bass21.dll" () As Long
Declare Function BASS_Pause Lib "bass21.dll" () As Long
Declare Function BASS_SetVolume Lib "bass21.dll" (ByVal volume As Long) As Long
Declare Function BASS_GetVolume Lib "bass21.dll" () As Long

Declare Function BASS_Set3DFactors Lib "bass21.dll" (ByVal distf As Single, ByVal rollf As Single, ByVal doppf As Single) As Long
Declare Function BASS_Get3DFactors Lib "bass21.dll" (ByRef distf As Single, ByRef rollf As Single, ByRef doppf As Single) As Long
Declare Function BASS_Set3DPosition Lib "bass21.dll" (ByRef pos As Any, ByRef vel As Any, ByRef front As Any, ByRef top As Any) As Long
Declare Function BASS_Get3DPosition Lib "bass21.dll" (ByRef pos As Any, ByRef vel As Any, ByRef front As Any, ByRef top As Any) As Long
Declare Function BASS_Apply3D Lib "bass21.dll" () As Long
Declare Function BASS_SetEAXParameters Lib "bass21.dll" (ByVal env As Long, ByVal vol As Single, ByVal decay As Single, ByVal damp As Single) As Long
Declare Function BASS_GetEAXParameters Lib "bass21.dll" (ByRef env As Long, ByRef vol As Single, ByRef decay As Single, ByRef damp As Single) As Long

Declare Function BASS_MusicLoad Lib "bass21.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long, ByVal freq As Long) As Long
Declare Sub BASS_MusicFree Lib "bass21.dll" (ByVal handle As Long)
Declare Function BASS_MusicGetName Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_MusicGetLength Lib "bass21.dll" (ByVal handle As Long, ByVal playlen As Long) As Long
Declare Function BASS_MusicSetAttribute Lib "bass21.dll" (ByVal handle As Long, ByVal attrib As Long, ByVal value As Long) As Long
Declare Function BASS_MusicGetAttribute Lib "bass21.dll" (ByVal handle As Long, ByVal attrib As Long) As Long

Declare Function BASS_SampleLoad Lib "bass21.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal max As Long, ByVal flags As Long) As Long
Declare Function BASS_SampleCreate Lib "bass21.dll" (ByVal length As Long, ByVal freq As Long, ByVal max As Long, ByVal flags As Long) As Long
Declare Function BASS_SampleCreateDone Lib "bass21.dll" () As Long
Declare Sub BASS_SampleFree Lib "bass21.dll" (ByVal handle As Long)
Declare Function BASS_SampleGetInfo Lib "bass21.dll" (ByVal handle As Long, ByRef info As BASS_SAMPLE) As Long
Declare Function BASS_SampleSetInfo Lib "bass21.dll" (ByVal handle As Long, ByRef info As BASS_SAMPLE) As Long
Declare Function BASS_SampleGetChannel Lib "bass21.dll" (ByVal handle As Long, ByVal onlynew As Long) As Long
Declare Function BASS_SampleStop Lib "bass21.dll" (ByVal handle As Long) As Long

Declare Function BASS_StreamCreate Lib "bass21.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_StreamCreateFile Lib "bass21.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Declare Function BASS_StreamCreateURL Lib "bass21.dll" (ByVal url As String, ByVal offset As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_StreamCreateFileUser Lib "bass21.dll" (ByVal buffered As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Sub BASS_StreamFree Lib "bass21.dll" (ByVal handle As Long)
Declare Function BASS_StreamGetLength Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_StreamGetTags Lib "bass21.dll" (ByVal handle As Long, ByVal tags As Long) As Long
Declare Function BASS_StreamGetFilePosition Lib "bass21.dll" (ByVal handle As Long, ByVal mode As Long) As Long

Declare Function BASS_RecordGetDeviceDescription Lib "bass21.dll" (ByVal device As Long) As Long
Declare Function BASS_RecordInit Lib "bass21.dll" (ByVal device As Long) As Long
Declare Function BASS_RecordSetDevice Lib "bass21.dll" (ByVal device As Long) As Long
Declare Function BASS_RecordGetDevice Lib "bass21.dll" () As Long
Declare Function BASS_RecordFree Lib "bass21.dll" () As Long
Declare Function BASS_RecordGetInfo Lib "bass21.dll" (ByRef info As BASS_RECORDINFO) As Long
Declare Function BASS_RecordGetInputName Lib "bass21.dll" (ByVal inputn As Long) As Long
Declare Function BASS_RecordSetInput Lib "bass21.dll" (ByVal inputn As Long, ByVal setting As Long) As Long
Declare Function BASS_RecordGetInput Lib "bass21.dll" (ByVal inputn As Long) As Long
Declare Function BASS_RecordStart Lib "bass21.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long

Private Declare Function BASS_ChannelBytes2Seconds64 Lib "bass21.dll" Alias "BASS_ChannelBytes2Seconds" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Single
Declare Function BASS_ChannelSeconds2Bytes Lib "bass21.dll" (ByVal handle As Long, ByVal pos As Single) As Long
Declare Function BASS_ChannelGetDevice Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelIsActive Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetInfo Lib "bass21.dll" (ByVal handle As Long, ByRef info As BASS_CHANNELINFO) As Long
Declare Function BASS_ChannelSetFlags Lib "bass21.dll" (ByVal handle As Long, ByVal flags As Long) As Long
Declare Function BASS_ChannelPreBuf Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelPlay Lib "bass21.dll" (ByVal handle As Long, ByVal restart As Long) As Long
Declare Function BASS_ChannelStop Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelPause Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelSetAttributes Lib "bass21.dll" (ByVal handle As Long, ByVal freq As Long, ByVal volume As Long, ByVal pan As Long) As Long
Declare Function BASS_ChannelGetAttributes Lib "bass21.dll" (ByVal handle As Long, ByRef freq As Long, ByRef volume As Long, ByRef pan As Long) As Long
Declare Function BASS_ChannelSlideAttributes Lib "bass21.dll" (ByVal handle As Long, ByVal freq As Long, ByVal volume As Long, ByVal pan As Long, ByVal time As Long) As Long
Declare Function BASS_ChannelIsSliding Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelSet3DAttributes Lib "bass21.dll" (ByVal handle As Long, ByVal mode As Long, ByVal min As Single, ByVal max As Single, ByVal iangle As Long, ByVal oangle As Long, ByVal outvol As Long) As Long
Declare Function BASS_ChannelGet3DAttributes Lib "bass21.dll" (ByVal handle As Long, ByRef mode As Long, ByRef min As Single, ByRef max As Single, ByRef iangle As Long, ByRef oangle As Long, ByRef outvol As Long) As Long
Declare Function BASS_ChannelSet3DPosition Lib "bass21.dll" (ByVal handle As Long, ByRef pos As Any, ByRef orient As Any, ByRef vel As Any) As Long
Declare Function BASS_ChannelGet3DPosition Lib "bass21.dll" (ByVal handle As Long, ByRef pos As Any, ByRef orient As Any, ByRef vel As Any) As Long
Private Declare Function BASS_ChannelSetPosition64 Lib "bass21.dll" Alias "BASS_ChannelSetPosition" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Long
Declare Function BASS_ChannelGetPosition Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetLevel Lib "bass21.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetData Lib "bass21.dll" (ByVal handle As Long, ByRef buffer As Any, ByVal length As Long) As Long
Private Declare Function BASS_ChannelSetSync64 Lib "bass21.dll" Alias "BASS_ChannelSetSync" (ByVal handle As Long, ByVal atype As Long, ByVal param As Long, ByVal paramhigh As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_ChannelRemoveSync Lib "bass21.dll" (ByVal handle As Long, ByVal sync As Long) As Long
Declare Function BASS_ChannelSetDSP Lib "bass21.dll" (ByVal handle As Long, ByVal proc As Long, ByVal user As Long, ByVal priority As Long) As Long
Declare Function BASS_ChannelRemoveDSP Lib "bass21.dll" (ByVal handle As Long, ByVal dsp As Long) As Long
Declare Function BASS_ChannelSetEAXMix Lib "bass21.dll" (ByVal handle As Long, ByVal mix As Single) As Long
Declare Function BASS_ChannelGetEAXMix Lib "bass21.dll" (ByVal handle As Long, ByRef mix As Single) As Long
Declare Function BASS_ChannelSetLink Lib "bass21.dll" (ByVal handle As Long, ByVal chan As Long) As Long
Declare Function BASS_ChannelRemoveLink Lib "bass21.dll" (ByVal handle As Long, ByVal chan As Long) As Long
Declare Function BASS_ChannelSetFX Lib "bass21.dll" (ByVal handle As Long, ByVal atype As Long, ByVal priority As Long) As Long
Declare Function BASS_ChannelRemoveFX Lib "bass21.dll" (ByVal handle As Long, ByVal fx As Long) As Long
Declare Function BASS_FXSetParameters Lib "bass21.dll" (ByVal handle As Long, ByRef par As Any) As Long
Declare Function BASS_FXGetParameters Lib "bass21.dll" (ByVal handle As Long, ByRef par As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Public Function BASS_SPEAKER_N(ByVal n As Long) As Long
BASS_SPEAKER_N = n * (2 ^ 24)
End Function

'*******************************************
' 32-bit wrappers for 64-bit BASS functions
'*******************************************
Function BASS_ChannelBytes2Seconds(ByVal handle As Long, ByVal pos As Long) As Single
BASS_ChannelBytes2Seconds = BASS_ChannelBytes2Seconds64(handle, pos, 0)
End Function

Function BASS_ChannelSetPosition(ByVal handle As Long, ByVal pos As Long) As Long
BASS_ChannelSetPosition = BASS_ChannelSetPosition64(handle, pos, 0)
End Function

Function BASS_ChannelSetSync(ByVal handle As Long, ByVal atype As Long, ByVal param As Long, ByVal proc As Long, ByVal user As Long) As Long
BASS_ChannelSetSync = BASS_ChannelSetSync64(handle, atype, param, 0, proc, user)
End Function


Function STREAMPROC(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    
    'CALLBACK FUNCTION !!!
    
    ' User stream callback function
    ' NOTE: A stream function should obviously be as quick
    ' as possible, other streams (and MOD musics) can't be mixed until it's finished.
    ' handle : The stream that needs writing
    ' buffer : Buffer to write the samples in
    ' length : Number of bytes to write
    ' user   : The 'user' parameter value given when calling BASS_StreamCreate
    ' RETURN : Number of bytes written. Set the BASS_STREAMPROC_END flag to end
    '          the stream.
    
End Function

Function STREAMFILEPROC(ByVal action As Long, ByVal param1 As Long, ByVal param2 As Long, ByVal user As Long) As Long
    
    'CALLBACK FUNCTION !!!
    
    ' User file stream callback function.
    ' action : The action to perform, one of BASS_FILE_xxx values.
    ' param1 : Depends on "action"
    ' param2 : Depends on "action"
    ' user   : The 'user' parameter value given when calling BASS_StreamCreate
    ' RETURN : Depends on "action"
    
End Function

Sub DOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    
    'CALLBACK FUNCTION !!!

    ' Internet stream download callback function.
    ' buffer : Buffer containing the downloaded data... NULL=end of download
    ' length : Number of bytes in the buffer
    ' user   : The 'user' parameter given when calling BASS_StreamCreateURL
    
End Sub

Sub SYNCPROC(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    
    'CALLBACK FUNCTION !!!
    
    'Similarly in here, write what to do when sync function
    'is called, i.e screen flash etc.
    
    ' NOTE: a sync callback function should be very quick as other
    ' syncs cannot be processed until it has finished.
    ' handle : The sync that has occured
    ' channel: Channel that the sync occured in
    ' data   : Additional data associated with the sync's occurance
    ' user   : The 'user' parameter given when calling BASS_ChannelSetSync */
    
End Sub

Sub DSPPROC(ByVal handle As Long, ByVal channel As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long)

    'CALLBACK FUNCTION !!!

    ' VB doesn't support pointers, so you should copy the buffer into an array,
    ' process it, and then copy it back into the buffer.

    ' DSP callback function. NOTE: A DSP function should obviously be as quick as
    ' possible... other DSP functions, streams and MOD musics can not be processed
    ' until it's finished.
    ' handle : The DSP handle
    ' channel: Channel that the DSP is being applied to
    ' buffer : Buffer to apply the DSP to
    ' length : Number of bytes in the buffer
    ' user   : The 'user' parameter given when calling BASS_ChannelSetDSP
    
End Sub

Function RECORDPROC(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long

    'CALLBACK FUNCTION !!!

    ' Recording callback function.
    ' handle : The recording handle
    ' buffer : Buffer containing the recorded samples
    ' length : Number of bytes
    ' user   : The 'user' parameter value given when calling BASS_RecordStart
    ' RETURN : BASSTRUE = continue recording, BASSFALSE = stop

End Function


Function BASS_GetDeviceDescriptionString(ByVal device As Long) As String
Dim pstring As Long
Dim sstring As String
On Error Resume Next
pstring = BASS_GetDeviceDescription(device)
If pstring Then
    sstring = VBStrFromAnsiPtr(pstring)
End If
BASS_GetDeviceDescriptionString = sstring
End Function

Public Function BASS_MusicGetNameString(ByVal handle As Long) As String
Dim pstring As Long
Dim sstring As String
On Error Resume Next
pstring = BASS_MusicGetName(handle)
If pstring Then
    sstring = VBStrFromAnsiPtr(pstring)
End If
BASS_MusicGetNameString = sstring
End Function

Function BASS_SetEAXPreset(Preset) As Long
' This function is a workaround, because VB doesn't support multiple comma seperated
' paramaters for each Global Const, simply pass the EAX_ENVIRONMENT_xxx value to this function
' instead of BASS_SetEAXParameters as you would do in C++
Select Case Preset
    Case EAX_ENVIRONMENT_GENERIC
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_GENERIC, 0.5, 1.493, 0.5)
    Case EAX_ENVIRONMENT_PADDEDCELL
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PADDEDCELL, 0.25, 0.1, 0)
    Case EAX_ENVIRONMENT_ROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ROOM, 0.417, 0.4, 0.666)
    Case EAX_ENVIRONMENT_BATHROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_BATHROOM, 0.653, 1.499, 0.166)
    Case EAX_ENVIRONMENT_LIVINGROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_LIVINGROOM, 0.208, 0.478, 0)
    Case EAX_ENVIRONMENT_STONEROOM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_STONEROOM, 0.5, 2.309, 0.888)
    Case EAX_ENVIRONMENT_AUDITORIUM
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_AUDITORIUM, 0.403, 4.279, 0.5)
    Case EAX_ENVIRONMENT_CONCERTHALL
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CONCERTHALL, 0.5, 3.961, 0.5)
    Case EAX_ENVIRONMENT_CAVE
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CAVE, 0.5, 2.886, 1.304)
    Case EAX_ENVIRONMENT_ARENA
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ARENA, 0.361, 7.284, 0.332)
    Case EAX_ENVIRONMENT_HANGAR
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_HANGAR, 0.5, 10, 0.3)
    Case EAX_ENVIRONMENT_CARPETEDHALLWAY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CARPETEDHALLWAY, 0.153, 0.259, 2)
    Case EAX_ENVIRONMENT_HALLWAY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_HALLWAY, 0.361, 1.493, 0)
    Case EAX_ENVIRONMENT_STONECORRIDOR
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_STONECORRIDOR, 0.444, 2.697, 0.638)
    Case EAX_ENVIRONMENT_ALLEY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_ALLEY, 0.25, 1.752, 0.776)
    Case EAX_ENVIRONMENT_FOREST
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_FOREST, 0.111, 3.145, 0.472)
    Case EAX_ENVIRONMENT_CITY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_CITY, 0.111, 2.767, 0.224)
    Case EAX_ENVIRONMENT_MOUNTAINS
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_MOUNTAINS, 0.194, 7.841, 0.472)
    Case EAX_ENVIRONMENT_QUARRY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_QUARRY, 1, 1.499, 0.5)
    Case EAX_ENVIRONMENT_PLAIN
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PLAIN, 0.097, 2.767, 0.224)
    Case EAX_ENVIRONMENT_PARKINGLOT
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PARKINGLOT, 0.208, 1.652, 1.5)
    Case EAX_ENVIRONMENT_SEWERPIPE
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_SEWERPIPE, 0.652, 2.886, 0.25)
    Case EAX_ENVIRONMENT_UNDERWATER
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_UNDERWATER, 1, 1.499, 0)
    Case EAX_ENVIRONMENT_DRUGGED
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_DRUGGED, 0.875, 8.392, 1.388)
    Case EAX_ENVIRONMENT_DIZZY
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_DIZZY, 0.139, 17.234, 0.666)
    Case EAX_ENVIRONMENT_PSYCHOTIC
        BASS_SetEAXPreset = BASS_SetEAXParameters(EAX_ENVIRONMENT_PSYCHOTIC, 0.486, 7.563, 0.806)
End Select
End Function

Public Function HiWord(lparam As Long) As Long
' This is the HIWORD of the lParam:
HiWord = lparam \ &H10000 And &HFFFF&
End Function
Public Function LoWord(lparam As Long) As Long
' This is the LOWORD of the lParam:
LoWord = lparam And &HFFFF&
End Function
Function MakeLong(LoWord As Long, HiWord As Long) As Long
'Replacement for the c++ Function MAKELONG
MakeLong = (LoWord And &HFFFF&) Or (HiWord * &H10000)
End Function

Public Function VBStrFromAnsiPtr(ByVal lpStr As Long) As String
Dim bStr() As Byte
Dim cChars As Long
On Error Resume Next
' Get the number of characters in the buffer
cChars = lstrlen(lpStr)

' Resize the byte array
ReDim bStr(0 To cChars - 1) As Byte

' Grab the ANSI buffer
Call CopyMemory(bStr(0), ByVal lpStr, cChars)

' Now convert to a VB Unicode string
VBStrFromAnsiPtr = StrConv(bStr, vbUnicode)
End Function
