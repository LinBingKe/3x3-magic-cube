Attribute VB_Name = "NativeMILDeclarations"
'******************************************************************************
'*
'* Filename     :  MILVB.BAS
'* Owner        :  Matrox Imaging dept.
'* Rev          :  $Revision:   5.0  $
'* Content      :  This file contains all the corrresponding C header file
'*                 to use MIL with Visual Basic.
'*                 Are included: MIL.BAS
'*                               MILBLOB.BAS
'*                               MILMEAS.BAS
'*                               MILPAT.BAS
'*                               MILOCR.BAS
'*                               MILPROTO.BAS
'*                               MILSETUP.BAS
'*                               GENESIS.BAS
'*                               METEOR.BAS
'*                               PULSAR.BAS
'*
'* COPYRIGHT (c) Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************


'******************************************************************************
'******************************************************************************
'******************************************************************************
'*
'* Filename     :  MIL.BAS
'* Owner        :  Matrox Imaging dept.
'* Rev          :  $Revision:   5.0  $
'* Content      :  This file contains the defines necessary to use the
'*                 Matrox Imaging Library (MIL 4.0 and up) "VB" user interface.
'*
'* Comments     :  Some defines may be here but not yet
'*                 implemented in the library
'*
'* COPYRIGHT (c) Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'*************************************************************************/

'/************************************************************************/
'/* Support for old defines name                                         */
'/************************************************************************/

'/************************************************************************/
'/* MIL HOST CURRENT VERSION  (Inquired by MappInquire)                  */
'/************************************************************************/
Global Const M_MIL_CURRENT_VERSION = 5#


'/************************************************************************/
'/* FUNCTION ARGUMENT DECLARATIONS                                       */
'/************************************************************************/


'/************************************************************************/
'/* general default parameters (may be bit encoded)                      */
'/************************************************************************/
Global Const M_NULL = &H0&
Global Const M_FALSE = 0&
Global Const M_NO = 0&
Global Const M_OFF = 0&
Global Const M_IN_PROGRESS = 0&
Global Const M_FINISHED = 1&
Global Const M_TRUE = 1&
Global Const M_YES = 1&
Global Const M_ON = 1&
Global Const M_WAIT = 1&
Global Const M_CREATE = M_YES
Global Const M_FREE = M_NO
Global Const M_DEFAULT = &H10000000
Global Const M_QUIET = &H8000000
Global Const M_VALID = &H1&
Global Const M_INVALID = -1&
Global Const M_CLEAR = &H1&
Global Const M_NO_CLEAR = &H2&
Global Const M_LUT_OFFSET = &H80000000
Global Const M_ENABLE = -9997&
Global Const M_DISABLE = -9999&

Global Const M_EXTENDED = &H80000000                                                 '// remove
Global Const M_EXTENDED_ATTRIBUTE = M_EXTENDED                                       '// remove

'/************************************************************************/
'/* buffer ID offset for defaults                                        */
'/************************************************************************/
Global Const M_UNSIGNED = &H0&
Global Const M_SIGNED = &H80000000
Global Const M_FLOAT = (&H40000000 Or M_SIGNED)
Global Const M_DOUBLE = (&H20000000 Or M_SIGNED)
Global Const M_SIZE_BIT_MASK = &HFF&
Global Const M_TYPE_MASK = &HFFFFFF00

'/************************************************************************/
'/* MMX related                                                          */
'/************************************************************************/
Global Const MMX_EXTRA_BYTES = 32

'/************************************************************************/
'/* Multi thread                                                         */
'/************************************************************************/
Global Const M_MULTI_THREAD = &H1&
Global Const M_HOST_THREAD = &H2&
Global Const M_MIL_THREAD = &H4&
Global Const M_STATE = &H8&
Global Const M_SIGNALED = &H10&
Global Const M_NOT_SIGNALED = &H20&
Global Const M_THREAD_SELECT = &H40&
Global Const M_THREAD_DETACH = &H80&
Global Const M_EVENT_WAIT = &H100&
Global Const M_EVENT_STATE = &H200&
Global Const M_EVENT_SET = &H400&
Global Const M_AUTO_RESET = &H2000&
Global Const M_MANUAL_RESET = &H4000&
                                                      
Global Const M_EVENT_ALLOC = 1700&
Global Const M_EVENT_FREE = 1701&
Global Const M_EVENT_SEND = 1702&
Global Const M_EVENT_CONTROL = 1703&
Global Const M_EVENT_SYNCHRONIZE = 1704&
Global Const M_THREAD_ALLOC = 1800&
Global Const M_THREAD_FREE = 1801&
Global Const M_THREAD_WAIT = 1802&
Global Const M_THREAD_CONTROL = 1803&
Global Const M_THREAD_MODE = 1804&
Global Const M_THREAD_IO_MODE = 1805&

'/************************************************************************/
'/* General Inquire/Control ...                                          */
'/************************************************************************/
Global Const M_OWNER_APPLICATION = 1000&
Global Const M_OWNER_SYSTEM = 1001&
Global Const M_SIZE_X = 1002&
Global Const M_SIZE_Y = 1003&
Global Const M_SIZE_Z = 1004&
Global Const M_SIZE_BAND = 1005&
Global Const M_SIZE_BAND_LUT = 1006&
Global Const M_SIZE_BIT = 1007&
Global Const M_TYPE = 1008&
Global Const M_NUMBER = 1009&
Global Const M_FORMAT = 1010&
Global Const M_FORMAT_SIZE = 1011&
Global Const M_INIT_FLAG = 1012&
Global Const M_ATTRIBUTE = 1013&
Global Const M_SIGN = 1014&
Global Const M_LUT_ID = 1015&
Global Const M_NATIVE_ID = 1016&
Global Const M_COLOR_MODE = 1018&
Global Const M_THREAD_PRIORITY = 1019&
Global Const M_NEED_UPDATE = 1020&
Global Const M_SURFACE = 1021&
Global Const M_WINDOW_DDRAW_SURFACE = 1022&
Global Const M_OWNER_SYSTEM_TYPE = 1023&
Global Const M_DISP_NATIVE_ID = 1024&
Global Const M_ENCODER = 1025&
Global Const M_ENCODER_MODE = 1026&
Global Const M_ENCODER_TYPE = 1027&
Global Const M_ENCODER_SYNC = 1028&


'/************************************************************************/
'/* MsysAlloc()                                                          */
'/************************************************************************/
'/* System type */
Global Const M_DEFAULT_HOST = &H10000001
Global Const M_SYSTEM_HOST_TYPE = 9&
Global Const M_SYSTEM_MMX_TYPE = 10&
Global Const M_SYSTEM_MAGIC_TYPE = 10&
Global Const M_SYSTEM_IP8_TYPE = 11&
Global Const M_SYSTEM_IMAGE_TYPE = 12&
Global Const M_SYSTEM_VGA_TYPE = 13&
Global Const M_SYSTEM_COMET_TYPE = 14&
Global Const M_SYSTEM_METEOR_TYPE = 15&
Global Const M_SYSTEM_PULSAR_TYPE = 16&
Global Const M_SYSTEM_GENESIS_TYPE = 17&
Global Const M_SYSTEM_CORONA_TYPE = 18&
Global Const M_SYSTEM_VIDCAP_TYPE = 19&
Global Const M_SYSTEM_METEOR_II_TYPE = 20&

'/* MsysAlloc() flags  */
Global Const M_COMPLETE = &H0&
Global Const M_PARTIAL = &H1&
Global Const M_WINDOWS = &H2&
Global Const M_DISP_WAIT_SELECT = &H4&
Global Const M_DISP_TEXT_SAVE = &H8&
Global Const M_USE_DMA_FOR_PROC_BUF = &H10&
Global Const M_USE_DMA_FOR_DISP_BUF = &H20&
Global Const M_USE_DMA_FOR_GRAB_BUF = &H40&
Global Const M_PRE_ALLOC_DMA_MEM = &H80&
'/* Reserve next 8 bits for DMA                 from   0x00000100L */
'/* Size at allocation                                 0x00000200L */
'/*                                                    0x00000400L */
'/*                                                    0x00000800L */
'/*                                                    0x00001000L */
'/*                                                    0x00002000L */
'/*                                                    0x00004000L */
'/*                                             to     0x00008000L */
Global Const M_DMA_MEM_MASK = &HFF00&
Global Const M_NO_INTERRUPT = &H10000
Global Const M_NO_FIELD_START_INTERRUPT = &H20000
Global Const M_DISP_NO_WAIT_SELECT = &H40000
Global Const M_NO_DDRAW = &H80000
Global Const M_EXTERNAL_CLK_TTL = &H100000
Global Const M_EXTERNAL_CLK_422 = &H200000
Global Const M_DDRAW = &H800000

Global Const M_DMA_BLOCK_SIZE = 64&

Global Const M_USE_DMA = (M_USE_DMA_FOR_PROC_BUF Or M_USE_DMA_FOR_DISP_BUF Or M_USE_DMA_FOR_GRAB_BUF)


'/************************************************************************/
'/* SysAlloc() in Mil Interpreter                                        */
'/************************************************************************/
Global Const M_SYSTEM_HOST_PTR = (M_SYSTEM_HOST_TYPE + 50&)
Global Const M_SYSTEM_MAGIC_PTR = (M_SYSTEM_MAGIC_TYPE + 50&)
Global Const M_SYSTEM_IP8_PTR = (M_SYSTEM_IP8_TYPE + 50&)
Global Const M_SYSTEM_IMAGE_PTR = (M_SYSTEM_IMAGE_TYPE + 50&)
Global Const M_SYSTEM_VGA_PTR = (M_SYSTEM_VGA_TYPE + 50&)
Global Const M_SYSTEM_COMET_PTR = (M_SYSTEM_COMET_TYPE + 50&)
Global Const M_SYSTEM_METEOR_PTR = (M_SYSTEM_METEOR_TYPE + 50&)
Global Const M_SYSTEM_PULSAR_PTR = (M_SYSTEM_PULSAR_TYPE + 50&)
Global Const M_SYSTEM_GENESIS_PTR = (M_SYSTEM_GENESIS_TYPE + 50&)
Global Const M_SYSTEM_CORONA_PTR = (M_SYSTEM_CORONA_TYPE + 50&)
Global Const M_SYSTEM_VIDCAP_PTR = (M_SYSTEM_VIDCAP_TYPE + 50&)
Global Const M_SYSTEM_METEOR_II_PTR = (M_SYSTEM_METEOR_II_TYPE + 50&)


'/************************************************************************/
'/* MsysInquire() / MsysControl() Types                                  */
'/************************************************************************/

Global Const M_SYSTEM_TYPE = 2000&
Global Const M_SYSTEM_TYPE_PTR = 2001&
Global Const M_DISPLAY_NUM = 2002&
Global Const M_DISPLAY_TYPE = 2003&
Global Const M_DIGITIZER_NUM = 2004&
Global Const M_DIGITIZER_TYPE = 2005&
Global Const M_PROCESSOR_NUM = 2006&
Global Const M_PROCESSOR_TYPE = 2007&
Global Const M_PROCESSING_SYSTEM = 2008&
Global Const M_PROCESSING_SYSTEM_TYPE = 2009&
Global Const M_TUNER_NUM = 2010&
Global Const M_TUNER_TYPE = 2011&
Global Const M_RGB_MODULE_NUM = 2012&
Global Const M_RGB_MODULE_TYPE = 2013&
Global Const M_BOARD_TYPE = 2014&
Global Const M_BOARD_REVISION = 2015&
Global Const M_DISPLAY_LIST = 2016&
Global Const M_WIN_MODE = 2017&
Global Const M_DUAL_SCREEN_MODE = 2018&
Global Const M_UNDERLAY_SURFACE_AVAILABLE = 2019&
Global Const M_UNDERLAY_SURFACE_PHYSICAL_ADDRESS = 2020&
Global Const M_MAX_TILE_SIZE = 2021&
Global Const M_MAX_TILE_SIZE_X = 2022&
Global Const M_MAX_TILE_SIZE_Y = 2023&
Global Const M_LOW_LEVEL_SYSTEM_ID = 2024&
Global Const M_NATIVE_THREAD_ID = 2026&
Global Const M_NATIVE_MODE_ENTER = 2027&
Global Const M_NATIVE_MODE_LEAVE = 2028&
Global Const M_PHYSICAL_ADDRESS_UNDERLAY = 2029&
Global Const M_PHYSICAL_ADDRESS_VGA = 2030&
Global Const M_PSEUDO_LIVE_GRAB_ON_MGA = 2031&
Global Const M_PSEUDO_LIVE_GRAB_WHEN_OVERLAPPED = 2032&
Global Const M_FORCE_PSEUDO_IN_NON_UNDERLAY_DISPLAYS = 2033&
Global Const M_LIVE_GRAB = 2034&
Global Const M_LIVE_GRAB_WHEN_DISPLAY_DOES_NOT_MATCH = 2035&
Global Const M_LIVE_GRAB_TRACK = 2036&
Global Const M_LIVE_GRAB_MOVE_UPDATE = 2037&
Global Const M_LIVE_GRAB_END_TRIGGER = 2038&
Global Const M_LIVE_GRAB_FAST_HALT = 2039&
Global Const M_STOP_LIVE_GRAB_WHEN_MENU = 2040&
Global Const M_STOP_LIVE_GRAB_WHEN_INACTIVE = 2041&
Global Const M_STOP_LIVE_GRAB_WHEN_DISABLED = 2042&
Global Const M_GRAB_BY_DISPLAY_CAPTURE = 2043&
Global Const M_ALLOC_BUF_RGB888_AS_RGB555 = 2044&
Global Const M_RGB555_BUFFER_ALLOCATION = 2045&
Global Const M_LAST_GRAB_IN_TRUE_BUFFER = 2046&
Global Const M_NO_GRAB_WHEN_NO_INPUT_SIGNAL = 2047&
Global Const M_PCI_LATENCY = 2048&
Global Const M_FAST_PCI_TO_MEM = 2049&
Global Const M_DCF_SUPPORTED = 2050&
Global Const M_DMA_ENABLE = 2051&
Global Const M_DMA_DISABLE = 2052&
Global Const M_DIB_ONLY = 2053&
Global Const M_DIB_OR_DDRAW = 2054&
Global Const M_FLIP_ONLY = 2055&
Global Const M_PRIMARY_DDRAW_SURFACE_PTR = 2056&
Global Const M_PRIMARY_DDRAW_SURFACE_MEM_PTR = 2057&
Global Const M_PRIMARY_DDRAW_SURFACE_PITCH = 2058&
Global Const M_PRIMARY_DDRAW_SURFACE_SIZE_X = 2059&
Global Const M_PRIMARY_DDRAW_SURFACE_SIZE_Y = 2060&
Global Const M_PRIMARY_DDRAW_SURFACE_SIZE_BITS = 2061&
Global Const M_INTERNAL_FORMAT_SIZE = 2062&
Global Const M_INTERNAL_FORMAT_ENUMERATION = 2063&
Global Const M_INTERNAL_FORMAT_CHECK = 2064&
Global Const M_DDRAW_AVAILABLE = 2065&
Global Const M_BOARD_CODE = 2066&
Global Const M_LIVE_GRAB_DDRAW = 2067&
Global Const M_THREAD_CONTEXT_PTR = 2068&
Global Const M_PSEUDO_LIVE_GRAB_NB_FRAMES = 2069&
Global Const M_PSEUDO_LIVE_GRAB_NB_FIELDS = 2070&
Global Const M_DISPLAY_DOUBLE_BUFFERING = 2071&
Global Const M_PSEUDO_LIVE_GRAB_TIME = 2072&
Global Const M_PCI_BRIDGE_LATENCY = 2073&
Global Const M_PSEUDO_LIVE_GRAB_DDRAW = 2074&
Global Const M_MULTI_DISP_IN_UNDERLAY = 2075&
Global Const M_MULTI_DISP_FOR_GRAB = 2076&
Global Const M_TIMEOUT = 2077&
Global Const M_AUTO_FLIP_FOR_TRUE_COLOR = 2078&
Global Const M_PCI_BRIDGE_HOST_WRITE_POSTING = 2079&
Global Const M_FAST_MEM_TO_VGA = 2080&
Global Const M_ERROR_ASYNCHRONOUS_LOG = 2081&
Global Const M_LIVE_GRAB_WHEN_NOT_VISIBLE = 2082&
Global Const M_USE_MMX = 2083&
Global Const M_OVERLAPPED_STRUC = 2085&
Global Const M_PHYSICAL_ADDRESS_VIA = 2086&
Global Const M_PCI_MGA_ID = 2087&
Global Const M_PCI_VIA_ID = 2088&
Global Const M_PCI_BRIDGE_ID = 2089&
Global Const M_NATIVE_SYSTEM_NUMBER = 2090&
Global Const M_NATIVE_NODE_NUMBER = 2091&
Global Const M_VIDCAP_WINDOW_HANDLE = 2092&

'// !!! MAP FOR OLD DEFINES
Global Const M_LIVE_VIDEO = M_LIVE_GRAB
Global Const M_LAST_GRAB_IN_ACTUAL_BUFFER = M_LAST_GRAB_IN_TRUE_BUFFER
Global Const M_SWITCH_TO_PSEUDO_WHEN_OVERLAPPED = M_PSEUDO_LIVE_GRAB_WHEN_OVERLAPPED
Global Const M_FORCE_PSEUDO_IN_NON_PULSAR_DISPLAYS = M_FORCE_PSEUDO_IN_NON_UNDERLAY_DISPLAYS
Global Const M_SYS_TYPE = M_SYSTEM_TYPE
Global Const M_SYS_TYPE_PTR = M_SYSTEM_TYPE_PTR
Global Const M_SYS_NUMBER = M_NUMBER
Global Const M_SYS_INIT_FLAG = M_INIT_FLAG
Global Const M_SYS_DISPLAY_NUM = M_DISPLAY_NUM
Global Const M_SYS_DISPLAY_TYPE = M_DISPLAY_TYPE
Global Const M_SYS_DIGITIZER_NUM = M_DIGITIZER_NUM
Global Const M_SYS_DIGITIZER_TYPE = M_DIGITIZER_TYPE
Global Const M_SYS_PROCESSOR_NUM = M_PROCESSOR_NUM
Global Const M_SYS_PROCESSOR_TYPE = M_PROCESSOR_TYPE
Global Const M_SYS_BOARD_TYPE = M_BOARD_TYPE
Global Const M_SYS_BOARD_REVISION = M_BOARD_REVISION
Global Const M_SYS_TUNER_NUM = M_TUNER_NUM
Global Const M_SYS_TUNER_TYPE = M_TUNER_TYPE
Global Const M_SYS_RGB_MODULE_NUM = M_RGB_MODULE_NUM
Global Const M_SYS_RGB_MODULE_TYPE = M_RGB_MODULE_TYPE
Global Const M_SYS_DISPLAY_LIST = M_DISPLAY_LIST
Global Const M_SYS_DUAL_SCREEN_MODE = M_DUAL_SCREEN_MODE
Global Const M_SYS_UNDERLAY_SURFACE_AVAILABLE = M_UNDERLAY_SURFACE_AVAILABLE
Global Const M_SYS_UNDERLAY_SURFACE_PHYSICAL_ADDRESS = M_UNDERLAY_SURFACE_PHYSICAL_ADDRESS
Global Const M_SYS_WIN_MODE = M_WIN_MODE
Global Const M_SYS_MAX_TILE_SIZE = M_MAX_TILE_SIZE
Global Const M_SYS_MAX_TILE_SIZE_X = M_MAX_TILE_SIZE_X
Global Const M_SYS_MAX_TILE_SIZE_Y = M_MAX_TILE_SIZE_Y
Global Const M_ON_BOARD_MEM_ADRS = M_PHYSICAL_ADDRESS_UNDERLAY
Global Const M_ON_BOARD_VGA_ADRS = M_PHYSICAL_ADDRESS_VGA


'/************************************************************************/
'/* MsysInquire() / MsysControl() Values                                 */
'/************************************************************************/


'/************************************************************************/
'/* MsysConfigAccess()                                                   */
'/************************************************************************/
Global Const M_PCI_CONFIGURATION_SPACE = 0&


'/************************************************************************/
'/* MdispAlloc() for VGA system                                          */
'/************************************************************************/
Global Const M_WINDOW_MAXIMIZE = &H8&
Global Const M_WINDOW_NO_MENUBAR = &H10&
Global Const M_WINDOW_NO_TITLEBAR = &H20&
Global Const M_WINDOW_NO_KEY = &H40&
Global Const M_WINDOW_USE_FORMAT = &H100&
Global Const M_PALETTE_MIL = &H0&
Global Const M_PALETTE_WINDOWS = &H200&
Global Const M_ZOOM_ENHANCED = &H0&
Global Const M_ZOOM_BASIC = &H400&
Global Const M_DISPLAY_8_BASIC = &H0&
Global Const M_DISPLAY_8_ENHANCED = &H800&
Global Const M_DISPLAY_24_ENHANCED = &H0&
Global Const M_DISPLAY_24_BASIC = &H1000&
Global Const M_DISPLAY_24_WINDOWS = &H2000&
Global Const M_DISPLAY_ENHANCED = (M_DISPLAY_8_ENHANCED + M_DISPLAY_24_ENHANCED)
Global Const M_DISPLAY_BASIC = (M_DISPLAY_8_BASIC + M_DISPLAY_24_BASIC)
Global Const M_DISPLAY_WINDOWS = (M_DISPLAY_8_BASIC + M_DISPLAY_24_WINDOWS)
Global Const M_WINDOW_NO_SYSBUTTON = &H4000&
Global Const M_WINDOW_NO_MINBUTTON = &H8000&
Global Const M_WINDOW_NO_MAXBUTTON = &H10000
Global Const M_COLORTABLE_INDEX = &H20000
Global Const M_COLORTABLE_RGB = &H0&
Global Const M_PALETTE_NOCOLLAPSE = &H100000
Global Const M_PALETTE_COLLAPSE = &H0&
Global Const M_USE_MEMORY_VCF = &H10&


'/************************************************************************/
'/* MdispAlloc() for Windowed system                                     */
'/************************************************************************/
Global Const M_WINDOWED = &H1000000
Global Const M_NON_WINDOWED = &H2000000
Global Const M_AUTOMATIC = &HFFFFFFFF
Global Const M_DEV0 = 0&
Global Const M_DEV1 = 1&
Global Const M_DEV2 = 2&
Global Const M_DEV3 = 3&
Global Const M_DEV4 = 4&
Global Const M_DEV5 = 5&
Global Const M_DEV6 = 6&
Global Const M_DEV7 = 7&
Global Const M_DEV8 = 8&
Global Const M_DEV9 = 9&
Global Const M_DEV10 = 10&
Global Const M_DEV11 = 11&
Global Const M_DEV12 = 12&
Global Const M_DEV13 = 13&
Global Const M_DEV14 = 14&
Global Const M_DEV15 = 15&
Global Const M_NODE0 = &H10000
Global Const M_NODE1 = &H20000
Global Const M_NODE2 = &H40000
Global Const M_NODE3 = &H80000
Global Const M_NODE4 = &H100000
Global Const M_NODE5 = &H200000
Global Const M_NODE6 = &H400000
Global Const M_NODE7 = &H800000
Global Const M_NODE8 = &H1000000
Global Const M_NODE9 = &H2000000
Global Const M_NODE10 = &H4000000
Global Const M_NODE11 = &H8000000
Global Const M_NODE12 = &H10000000
Global Const M_NODE13 = &H20000000
Global Const M_NODE14 = &H40000000
Global Const M_NODE15 = &H80000000


'/************************************************************************/
'/* MdispInquire() / MdispControl() Types                                */
'/************************************************************************/

Global Const M_PAN_X = 3000&
Global Const M_PAN_Y = 3001&
Global Const M_ZOOM_X = 3002&
Global Const M_ZOOM_Y = 3003&
Global Const M_HARDWARE_PAN = 3004&
Global Const M_HARDWARE_ZOOM = 3005&
Global Const M_SELECTED = 3006&
Global Const M_KEY_MODE = 3007&
Global Const M_KEY_CONDITION = 3008&
Global Const M_KEY_MASK = 3009&
Global Const M_KEY_COLOR = 3010&
Global Const M_KEY_SUPPORTED = 3011&
Global Const M_VGA_BUF_ID = 3012&
Global Const M_WINDOW_BUF_WRITE = 3013&
Global Const M_WINDOW_BUF_ID = 3014&
Global Const M_WINDOW_OVR_BUF_ID = 3015&
Global Const M_WINDOW_OVR_WRITE = 3016&
Global Const M_WINDOW_OVR_DISP_ID = 3017&
Global Const M_INTERPOLATION_MODE = 3018&
Global Const M_HOOK_OFFSET = 3019&
Global Const M_FRAME_START_HANDLER_PTR = 3020&
Global Const M_FRAME_START_HANDLER_USER_PTR = 3021&
Global Const M_WINDOW_OVR_LUT = 3022&
Global Const M_WINDOW_OVR_SHOW = 3023&
Global Const M_WINDOW_DISPLAY_SETTINGS = 3024&
Global Const M_WINDOW_OVR_LUT_REMAP = 3025&
Global Const M_WINDOW_AUTO_ACTIVATION_FOR_DDRAW = 3026&
Global Const M_DISPLAY_16_TO_8 = 3027&
Global Const M_DISPLAY_16_TO_8_SHIFT = 3028&
Global Const M_DISPLAY_MODE = 3029&
Global Const M_WINDOW_OVR_FLICKER = 3031&

Global Const M_WINDOW_ZOOM = 3051&
Global Const M_WINDOW_RESIZE = 3052&
Global Const M_WINDOW_OVERLAP = 3053&
Global Const M_WINDOW_SCROLLBAR = 3054&
Global Const M_WINDOW_UPDATE = 3055&
Global Const M_WINDOW_PROTECT_AREA = 3056&
Global Const M_WINDOW_TITLE_BAR = 3057&
Global Const M_WINDOW_MENU_BAR = 3058&
Global Const M_WINDOW_TITLE_BAR_CHANGE = 3059&
Global Const M_WINDOW_MENU_BAR_CHANGE = 3060&
Global Const M_WINDOW_MOVE = 3061&
Global Const M_WINDOW_SYSBUTTON = 3062&
Global Const M_WINDOW_MINBUTTON = 3063&
Global Const M_WINDOW_MAXBUTTON = 3064&
Global Const M_WINDOW_COLOR = 3065&
Global Const M_WINDOW_COLOR_CHANGE = 3066&
Global Const M_WINDOW_PALETTE = 3067&
Global Const M_WINDOW_PALETTE_WINDOWS = 3068&
Global Const M_WINDOW_PALETTE_NOCOLLAPSE = 3069&
Global Const M_WINDOW_PALETTE_BACKGROUND = 3070&
Global Const M_WINDOW_PALETTE_AUTO = 3071&
Global Const M_WINDOW_ERASE_BACKGROUND = 3072&
Global Const M_WINDOW_UPDATE_AUTO_ON_CONTROL = 3073&
Global Const M_WINDOW_UPDATE_WITH_SEND_MESSAGE = 3074&
Global Const M_WINDOW_SNAP_X = 3075&
Global Const M_WINDOW_SNAP_Y = 3076&
Global Const M_WINDOW_UPDATE_REGION = 3077&
Global Const M_WINDOW_UPDATE_ONLY_INVALID_BORDER = 3078&
Global Const M_WINDOW_UPDATE_KEEP_PALETTE_ALIVE = 3079&
Global Const M_WINDOW_UPDATE_ADD_BEGINPAINT = 3080&
Global Const M_WINDOW_UPDATE_ON_PAINT = 3081&
Global Const M_WINDOW_UPDATE_MANUAL = 3082&
Global Const M_WINDOW_PAINT = 3083&
Global Const M_WINDOW_ACTIVATE_DELAY = 3084&
Global Const M_WINDOW_CLIP_IN_CLIENT = 3085&
Global Const M_WINDOW_SYNC_SELECT = 3087&
Global Const M_WINDOW_INITIAL_POSITION_X = 3088&
Global Const M_WINDOW_INITIAL_POSITION_Y = 3089&
Global Const M_WINDOW_BENCHMARK_IN_DEBUG = 3090&
Global Const M_WINDOW_RANGE = 3091&
Global Const M_WINDOW_OVR_BUFFER_ALIVE = 3092&
Global Const M_WINDOW_OVR_BUFFER_PTR = 3093&
Global Const M_WINDOW_OVR_FLICKER_FREE_ALIVE = 3094&
Global Const M_WINDOW_OVR_FLICKER_FREE_PTR = 3095&
Global Const M_WINDOW_OVR_DESTRUCTIVE = 3096&
Global Const M_WINDOW_OVR_KEYER_PTR = 3097&
Global Const M_WINDOW_MANUAL_OVR_ADD = 3098&
Global Const M_WINDOW_MANUAL_FLICKER_COPY = 3099&
Global Const M_WINDOW_MANUAL_OVR_ADD_FLICKER_COPY = 3100&
Global Const M_WINDOW_USE_SUBCLASS_TRACKING = 3101&
Global Const M_WINDOW_USE_SYSTEMHOOK_TRACKING = 3102&
Global Const M_WINDOW_ATTRIBUTE_FOR_OVERLAY = 3103&
Global Const M_WINDOW_ATTRIBUTE_FOR_FLICKER = 3104&
Global Const M_WINDOW_MASK_FOR_OVERLAY_VERIFICATION = 3105&
Global Const M_WINDOW_MASK_FOR_FLICKER_VERIFICATION = 3106&
Global Const M_DESKTOP_CHANGE = 3107&
Global Const M_WINDOW_HOOK_BLOCKING_SERIALIZATION = 3108&
Global Const M_WINDOW_ATTRIBUTE_FOR_BUFFER = 3109&
Global Const M_WINDOW_HANDLE = 3110&
Global Const M_WINDOW_OFFSET_X = 3111&
Global Const M_WINDOW_OFFSET_Y = 3112&
Global Const M_WINDOW_SIZE_X = 3113&
Global Const M_WINDOW_SIZE_Y = 3114&
Global Const M_WINDOW_PAN_X = 3115&
Global Const M_WINDOW_PAN_Y = 3116&
Global Const M_WINDOW_ZOOM_X = 3117&
Global Const M_WINDOW_ZOOM_Y = 3118&
Global Const M_WINDOW_TITLE_NAME = 3119&
Global Const M_HOOK_MODIFIED_DIB_PTR = 3120&
Global Const M_WINDOW_USE_SYSTEMHOOK_TRACKING_ACTIVE = 3121&
'/* Reserve next 2 values                       from   3121L*/
'/*                                             to     3122L*/
Global Const M_HOOK_MODIFIED_DIB_USER_PTR = 3123&
'/* Reserve next 2 values                       from   3124L*/
'/*                                             to     3125L*/
Global Const M_HOOK_MODIFIED_WINDOW_PTR = 3126&
'/* Reserve next 2 values                       from   3127L*/
'/*                                             to     3128L*/
Global Const M_HOOK_MODIFIED_WINDOW_USER_PTR = 3129&
'/* Reserve next 2 values                       from   3130L*/
'/*                                             to     3131L*/
Global Const M_HOOK_MESSAGE_LOOP_PTR = 3132&
'/* Reserve next 2 values                       from   3133L*/
'/*                                             to     3134L*/
Global Const M_HOOK_MESSAGE_LOOP_USER_PTR = 3135&
'/* Reserve next 2 values                       from   3136L*/
'/*                                             to     3137L*/
Global Const M_WINDOW_APPFRAME_HANDLE = 3138&
Global Const M_WINDOW_MDICLIENT_HANDLE = 3139&
Global Const M_WINDOW_MDIFRAME_HANDLE = 3140&
Global Const M_VISIBLE_OFFSET_X = 3141&
Global Const M_VISIBLE_OFFSET_Y = 3142&
Global Const M_VISIBLE_SIZE_X = 3145&
Global Const M_VISIBLE_SIZE_Y = 3146&
Global Const M_WINDOW_DIB_HANDLE = 3147&
Global Const M_WINDOW_DISPLAY_DIB_HANDLE = 3148&
Global Const M_WINDOW_ACTIVE = 3149&
Global Const M_WINDOW_ENABLE = 3150&
Global Const M_PALETTE_HANDLE = 3151&
Global Const M_WINDOW_THREAD_HANDLE = 3152&
Global Const M_WINDOW_THREAD_ID = 3153&
Global Const M_WINDOW_DIB_HEADER = 3154&
Global Const M_WINDOW_KEYBOARD_USE = 3155&
Global Const M_WINDOW_CLIP_LIST_SIZE = 3156&
Global Const M_WINDOW_CLIP_LIST = 3157&
Global Const M_WINDOW_CLIP_LIST_ACCESS = 3158&
Global Const M_FRAME_START_TRIGGER_MODE = 3159&
Global Const M_FRAME_START_TRIGGER = 3160&
Global Const M_WINDOW_DIB = 3161&
Global Const M_WINDOW_MAP_BUFFER = 3162&
Global Const M_WINDOW_OVR_COPY = 3163&
Global Const M_WINDOW_UPDATE_EXCLUDE_RECTANGLE = 3164&
Global Const M_WINDOW_SYNC_UPDATE = 3165&
Global Const M_WINDOW_TITLE_NAME_SIZE = 3166&
Global Const M_WINDOW_DRIVER_SIZE_BIT = 3167&
Global Const M_WINDOW_SYNC_UPDATE_WHEN_HOOK_BLOCKED = 3168&
Global Const M_WINDOW_CLIP_LIST_BLOCKING_SERIALIZATION = 3169&
Global Const M_DESKTOP_LOCK_TIMEOUT = 3170&
Global Const M_WINDOW_PALETTE_MESSAGES = 3171&
Global Const M_WINDOW_PAINT_MESSAGES = 3172&
Global Const M_WINDOW_COMMAND_PROMPT_FULL_DRAG = 3173&
Global Const M_WINDOW_DISPLAY_MODE = 3174&
Global Const M_WINDOW_AUTO_UPDATE = 3073&
Global Const M_WINDOW_UPDATE_WITH_MESSAGE = 3074&
Global Const M_WINDOW_UPDATE_USE_BEGINPAINT = 3080&
Global Const M_WINDOW_UPDATE_USE_ERASEBKGND = 3081&
Global Const M_WINDOW_MANUAL_UPDATE = 3082&



'// !!! MAP FOR OLD DEFINES
Global Const M_DISP_LUT = M_LUT_ID
Global Const M_DISP_NUMBER = M_NUMBER
Global Const M_DISP_FORMAT = M_FORMAT
Global Const M_DISP_INIT_FLAG = M_INIT_FLAG
Global Const M_DISP_PAN_X = M_PAN_X
Global Const M_DISP_PAN_Y = M_PAN_Y
Global Const M_DISP_ZOOM_X = M_ZOOM_X
Global Const M_DISP_ZOOM_Y = M_ZOOM_Y
Global Const M_DISP_HARDWARE_PAN = M_HARDWARE_PAN
Global Const M_DISP_HARDWARE_ZOOM = M_HARDWARE_ZOOM
Global Const M_DISP_KEY_MODE = M_KEY_MODE
Global Const M_DISP_KEY_CONDITION = M_KEY_CONDITION
Global Const M_DISP_KEY_MASK = M_KEY_MASK
Global Const M_DISP_KEY_COLOR = M_KEY_COLOR
Global Const M_DISP_16_TO_8 = M_DISPLAY_16_TO_8
Global Const M_DISP_16_TO_8_SHIFT = M_DISPLAY_16_TO_8_SHIFT
Global Const M_DISP_MODE = M_DISPLAY_MODE
Global Const M_DISP_THREAD_PRIORITY = M_THREAD_PRIORITY
Global Const M_DISP_INTERPOLATION_MODE = M_INTERPOLATION_MODE
Global Const M_DISP_HOOK_OFFSET = M_HOOK_OFFSET
Global Const M_DISP_VGA_BUF_ID = M_VGA_BUF_ID
Global Const M_DISP_OVR_WRITE = M_WINDOW_OVR_WRITE
Global Const M_DISP_OVR_BUF_ID = M_WINDOW_OVR_BUF_ID
Global Const M_DISP_BUF_WRITE = M_WINDOW_BUF_WRITE
Global Const M_DISP_BUF_ID = M_WINDOW_BUF_ID
Global Const M_DISP_WINDOW_OVR_BUF_ID = M_WINDOW_OVR_BUF_ID
Global Const M_DISP_WINDOW_OVR_WRITE = M_WINDOW_OVR_WRITE
Global Const M_DISP_VGA_DISPLAY_ID = M_WINDOW_OVR_DISP_ID
Global Const M_DISP_KEY_SUPPORTED = M_KEY_SUPPORTED

Global Const M_DISP_WINDOW_ZOOM = M_WINDOW_ZOOM
Global Const M_DISP_WINDOW_RESIZE = M_WINDOW_RESIZE
Global Const M_DISP_WINDOW_OVERLAP = M_WINDOW_OVERLAP
Global Const M_DISP_WINDOW_SCROLLBAR = M_WINDOW_SCROLLBAR
Global Const M_DISP_WINDOW_UPDATE = M_WINDOW_UPDATE
Global Const M_DISP_WINDOW_PROTECT_AREA = M_WINDOW_PROTECT_AREA
Global Const M_DISP_WINDOW_TITLE_BAR = M_WINDOW_TITLE_BAR
Global Const M_DISP_WINDOW_MENU_BAR = M_WINDOW_MENU_BAR
Global Const M_DISP_WINDOW_TITLE_BAR_CHANGE = M_WINDOW_TITLE_BAR_CHANGE
Global Const M_DISP_WINDOW_MENU_BAR_CHANGE = M_WINDOW_MENU_BAR_CHANGE
Global Const M_DISP_WINDOW_MOVE = M_WINDOW_MOVE
Global Const M_DISP_WINDOW_SYSBUTTON = M_WINDOW_SYSBUTTON
Global Const M_DISP_WINDOW_MINBUTTON = M_WINDOW_MINBUTTON
Global Const M_DISP_WINDOW_MAXBUTTON = M_WINDOW_MAXBUTTON
Global Const M_DISP_WINDOW_COLOR = M_WINDOW_COLOR
Global Const M_DISP_WINDOW_COLOR_CHANGE = M_WINDOW_COLOR_CHANGE
Global Const M_DISP_WINDOW_PALETTE = M_WINDOW_PALETTE
Global Const M_DISP_WINDOW_PALETTE_WINDOWS = M_WINDOW_PALETTE_WINDOWS
Global Const M_DISP_WINDOW_PALETTE_NOCOLLAPSE = M_WINDOW_PALETTE_NOCOLLAPSE
Global Const M_DISP_WINDOW_PALETTE_BACKGROUND = M_WINDOW_PALETTE_BACKGROUND
Global Const M_DISP_WINDOW_PALETTE_AUTO = M_WINDOW_PALETTE_AUTO
Global Const M_DISP_WINDOW_ERASE_BACKGROUND = M_WINDOW_ERASE_BACKGROUND
Global Const M_DISP_WINDOW_AUTO_UPDATE = M_WINDOW_AUTO_UPDATE
Global Const M_DISP_WINDOW_UPDATE_WITH_MESSAGE = M_WINDOW_UPDATE_WITH_MESSAGE
Global Const M_DISP_WINDOW_SNAP_X = M_WINDOW_SNAP_X
Global Const M_DISP_WINDOW_SNAP_Y = M_WINDOW_SNAP_Y
Global Const M_DISP_WINDOW_UPDATE_REGION = M_WINDOW_UPDATE_REGION
Global Const M_DISP_WINDOW_UPDATE_ONLY_INVALID_BORDER = M_WINDOW_UPDATE_ONLY_INVALID_BORDER
Global Const M_DISP_WINDOW_UPDATE_KEEP_PALETTE_ALIVE = M_WINDOW_UPDATE_KEEP_PALETTE_ALIVE
Global Const M_DISP_WINDOW_UPDATE_USE_BEGINPAINT = M_WINDOW_UPDATE_USE_BEGINPAINT
Global Const M_DISP_WINDOW_UPDATE_USE_ERASEBKGND = M_WINDOW_UPDATE_USE_ERASEBKGND
Global Const M_DISP_WINDOW_MANUAL_UPDATE = M_WINDOW_MANUAL_UPDATE
Global Const M_DISP_WINDOW_PAINT = M_WINDOW_PAINT
Global Const M_DISP_WINDOW_ACTIVATE_DELAY = M_WINDOW_ACTIVATE_DELAY
Global Const M_DISP_WINDOW_CLIP_IN_CLIENT = M_WINDOW_CLIP_IN_CLIENT
Global Const M_DISP_WINDOW_SYNC_SELECT = M_WINDOW_SYNC_SELECT
Global Const M_DISP_WINDOW_INITIAL_POSITION_X = M_WINDOW_INITIAL_POSITION_X
Global Const M_DISP_WINDOW_INITIAL_POSITION_Y = M_WINDOW_INITIAL_POSITION_Y
Global Const M_DISP_WINDOW_BENCHMARK_IN_DEBUG = M_WINDOW_BENCHMARK_IN_DEBUG
Global Const M_DISP_WINDOW_RANGE = M_WINDOW_RANGE
Global Const M_DISP_WINDOW_OVR_BUFFER_ALIVE = M_WINDOW_OVR_BUFFER_ALIVE
Global Const M_DISP_WINDOW_OVR_BUFFER_PTR = M_WINDOW_OVR_BUFFER_PTR
Global Const M_DISP_WINDOW_OVR_FLICKER_FREE_ALIVE = M_WINDOW_OVR_FLICKER_FREE_ALIVE
Global Const M_DISP_WINDOW_OVR_FLICKER_FREE_PTR = M_WINDOW_OVR_FLICKER_FREE_PTR
Global Const M_DISP_WINDOW_OVR_DESTRUCTIVE = M_WINDOW_OVR_DESTRUCTIVE
Global Const M_DISP_WINDOW_OVR_KEYER_PTR = M_WINDOW_OVR_KEYER_PTR
Global Const M_DISP_WINDOW_MANUAL_OVR_ADD = M_WINDOW_MANUAL_OVR_ADD
Global Const M_DISP_WINDOW_MANUAL_FLICKER_COPY = M_WINDOW_MANUAL_FLICKER_COPY
Global Const M_DISP_WINDOW_MANUAL_OVR_ADD_FLICKER_COPY = M_WINDOW_MANUAL_OVR_ADD_FLICKER_COPY
Global Const M_DISP_WINDOW_USE_SUBCLASS_TRACKING = M_WINDOW_USE_SUBCLASS_TRACKING
Global Const M_DISP_WINDOW_USE_SYSTEMHOOK_TRACKING = M_WINDOW_USE_SYSTEMHOOK_TRACKING
Global Const M_DISP_WINDOW_ATTRIBUTE_FOR_OVERLAY = M_WINDOW_ATTRIBUTE_FOR_OVERLAY
Global Const M_DISP_WINDOW_ATTRIBUTE_FOR_FLICKER = M_WINDOW_ATTRIBUTE_FOR_FLICKER
Global Const M_DISP_WINDOW_MASK_FOR_OVERLAY_VERIFICATION = M_WINDOW_MASK_FOR_OVERLAY_VERIFICATION
Global Const M_DISP_WINDOW_MASK_FOR_FLICKER_VERIFICATION = M_WINDOW_MASK_FOR_FLICKER_VERIFICATION
Global Const M_DISP_DESKTOP_CHANGE = M_DESKTOP_CHANGE

Global Const M_DISP_WINDOW_HANDLE = M_WINDOW_HANDLE
Global Const M_DISP_WINDOW_OFFSET_X = M_WINDOW_OFFSET_X
Global Const M_DISP_WINDOW_OFFSET_Y = M_WINDOW_OFFSET_Y
Global Const M_DISP_WINDOW_SIZE_X = M_WINDOW_SIZE_X
Global Const M_DISP_WINDOW_SIZE_Y = M_WINDOW_SIZE_Y
Global Const M_DISP_WINDOW_PAN_X = M_WINDOW_PAN_X
Global Const M_DISP_WINDOW_PAN_Y = M_WINDOW_PAN_Y
Global Const M_DISP_WINDOW_ZOOM_X = M_WINDOW_ZOOM_X
Global Const M_DISP_WINDOW_ZOOM_Y = M_WINDOW_ZOOM_Y
Global Const M_DISP_WINDOW_TITLE_NAME = M_WINDOW_TITLE_NAME
Global Const M_DISP_HOOK_MODIFIED_DIB_PTR = M_HOOK_MODIFIED_DIB_PTR
Global Const M_DISP_HOOK_MODIFIED_DIB_USER_PTR = M_HOOK_MODIFIED_DIB_USER_PTR
Global Const M_DISP_HOOK_MODIFIED_WINDOW_PTR = M_HOOK_MODIFIED_WINDOW_PTR
Global Const M_DISP_HOOK_MODIFIED_WINDOW_USER_PTR = M_HOOK_MODIFIED_WINDOW_USER_PTR
Global Const M_DISP_HOOK_MESSAGE_LOOP_PTR = M_HOOK_MESSAGE_LOOP_PTR
Global Const M_DISP_HOOK_MESSAGE_LOOP_USER_PTR = M_HOOK_MESSAGE_LOOP_USER_PTR
Global Const M_DISP_WINDOW_APPFRAME_HANDLE = M_WINDOW_APPFRAME_HANDLE
Global Const M_DISP_WINDOW_MDICLIENT_HANDLE = M_WINDOW_MDICLIENT_HANDLE
Global Const M_DISP_WINDOW_MDIFRAME_HANDLE = M_WINDOW_MDIFRAME_HANDLE
Global Const M_DISP_VISIBLE_OFFSET_X = M_VISIBLE_OFFSET_X
Global Const M_DISP_VISIBLE_OFFSET_Y = M_VISIBLE_OFFSET_Y
Global Const M_DISP_VISIBLE_SIZE_X = M_VISIBLE_SIZE_X
Global Const M_DISP_VISIBLE_SIZE_Y = M_VISIBLE_SIZE_Y
Global Const M_DISP_WINDOW_DIB_HANDLE = M_WINDOW_DIB_HANDLE
Global Const M_DISP_WINDOW_DISPLAY_DIB_HANDLE = M_WINDOW_DISPLAY_DIB_HANDLE
Global Const M_DISP_WINDOW_ACTIVE = M_WINDOW_ACTIVE
Global Const M_DISP_WINDOW_ENABLE = M_WINDOW_ENABLE
Global Const M_DISP_PALETTE_HANDLE = M_PALETTE_HANDLE
Global Const M_DISP_WINDOW_THREAD_HANDLE = M_WINDOW_THREAD_HANDLE
Global Const M_DISP_WINDOW_THREAD_ID = M_WINDOW_THREAD_ID
Global Const M_DISP_WINDOW_DIB = M_WINDOW_DIB
Global Const M_DISP_WINDOW_CLIP_LIST_SIZE = M_WINDOW_CLIP_LIST_SIZE
Global Const M_DISP_WINDOW_CLIP_LIST = M_WINDOW_CLIP_LIST
Global Const M_DISP_WINDOW_CLIP_LIST_ACCESS = M_WINDOW_CLIP_LIST_ACCESS

Global Const M_DISP_WINDOW_CHANGE_TITLE_BAR = M_DISP_WINDOW_TITLE_BAR_CHANGE
Global Const M_DISP_WINDOW_CHANGE_MENU_BAR = M_DISP_WINDOW_MENU_BAR_CHANGE
Global Const M_DISP_WINDOW_CHANGE_COLOR = M_DISP_WINDOW_COLOR_CHANGE
Global Const M_DISP_WINDOW_DO_PAINT = M_DISP_WINDOW_PAINT
Global Const M_DISP_SELECT = M_SELECTED

                                                      
'/************************************************************************/
'/* MdispControl() / MdispInquire() Values                               */
'/************************************************************************/
                                                        
Global Const M_FULL_SIZE = 0&
Global Const M_NORMAL_SIZE = 1&

Global Const M_BENCHMARK_IN_DEBUG_ON = &H1&
Global Const M_BENCHMARK_IN_DEBUG_OFF = 0
Global Const M_BENCHMARK_IN_DEBUG_TRACE = &H2&
Global Const M_BENCHMARK_IN_DEBUG_NOTRACE = 0
Global Const M_BENCHMARK_IN_DEBUG_ALLSIZE = &H4&
Global Const M_BENCHMARK_IN_DEBUG_CSTSIZE = 0

Global Const M_DISPLAY_SCAN_LINE_START = &H0&
Global Const M_DISPLAY_SCAN_LINE_END = &HFFFFFFFF

Global Const M_INFINITE = &HFFFFFFFF
Global Const M_SLAVE = 0&
Global Const M_MASTER = 1&
                                                        
'/************************************************************************/
'/* MdispLut()                                                           */
'/************************************************************************/
Global Const M_PSEUDO = (M_LUT_OFFSET + 8&)
                                                        
                                                        
'/************************************************************************/
'/* MdispHook()                                                          */
'/************************************************************************/

'/* Defines for hook to modification to bitmap and window */
Global Const M_NOT_MODIFIED = 0                 '/* No changed at all              */
Global Const M_MODIFIED_LUT = 1                 '/* Disp lut is changed            */
Global Const M_MODIFIED_DIB = 2                 '/* Disp buffer data is changed    */
Global Const M_MODIFIED_ZOOM = 3                '/* Disp is zoomed                 */
Global Const M_MODIFIED_PAN = 4                 '/* Disp is panned                 */
Global Const M_MODIFIED_DIB_CREATION = 5        '/* Disp receives a new buffer ID  */
Global Const M_MODIFIED_DIB_DESTRUCTION = 6     '/* Disp receives a buffer ID 0    */
Global Const M_MODIFIED_WINDOW_CREATION = 7     '/* Wnd is created                 */
Global Const M_MODIFIED_WINDOW_DESTRUCTION = 8  '/* Wnd is destroyed               */
Global Const M_MODIFIED_WINDOW_LOCATION = 9     '/* Wnd size is changed            */
Global Const M_MODIFIED_WINDOW_OVERLAP = 11     '/* Wnd overlap is changed         */
Global Const M_MODIFIED_WINDOW_ICONIZED = 12    '/* Wnd is changed to iconic state */
Global Const M_MODIFIED_WINDOW_ZOOM = 13        '/* Wnd is zoomed                  */
Global Const M_MODIFIED_WINDOW_PAN = 14         '/* Wnd is panned                  */
Global Const M_MODIFIED_WINDOW_MENU = 15        '/* Wnd menu pulled-down           */
Global Const M_MODIFIED_WINDOW_PAINT = 16       '/* Wnd is painted with image      */
Global Const M_MODIFIED_WINDOW_ACTIVE = 17      '/* Wnd activation state changed   */
Global Const M_MODIFIED_WINDOW_ENABLE = 18      '/* Wnd enable state changed       */
Global Const M_MODIFIED_WINDOW_CLIP_LIST = 19   '/* Wnd clip list changed          */

'/* M_MODIFIED_WINDOW_MENU modification hook defines */
Global Const M_MODIFIED_SYS_MENU = &H100&
Global Const M_MODIFIED_APP_MENU = &H200&
Global Const M_MODIFIED_USER_APP_MENU = &H10000
Global Const M_MODIFIED_RESTORE_MENUITEM = &H1&
Global Const M_MODIFIED_MOVE_MENUITEM = &H2&
Global Const M_MODIFIED_SIZE_MENUITEM = &H3&
Global Const M_MODIFIED_MINIMIZE_MENUITEM = &H4&
Global Const M_MODIFIED_MAXIMIZE_MENUITEM = &H5&
Global Const M_MODIFIED_CLOSE_MENUITEM = &H6&
Global Const M_MODIFIED_TASKLIST_MENUITEM = &H7&
Global Const M_MODIFIED_MENUBAR_MENUITEM = &H8&
Global Const M_MODIFIED_TITLEOFF_MENUITEM = &H9&
Global Const M_MODIFIED_ZOOMIN_MENUITEM = &HA&
Global Const M_MODIFIED_ZOOMOUT_MENUITEM = &HB&
Global Const M_MODIFIED_NOZOOM_MENUITEM = &HC&

'/* M_MODIFIED_WINDOW_ACTIVE modification hook defines */
'/* M_MODIFIED_WINDOW_ENABLE modification hook defines */
Global Const M_MODIFIED_STATE_FROM_WINDOW = 0
Global Const M_MODIFIED_STATE_FROM_PARENT = &H10&
Global Const M_MODIFIED_OFF = 0
Global Const M_MODIFIED_ON = &H1&
                                                        
'/* M_MODIFIED_WINDOW_CLIP_LIST modification hook defines */
'/* M_MODIFIED_WINDOW_CLIP_LIST modification hook defines */
Global Const M_MODIFIED_ACCESS_RECTANGULAR_OFF = 0
Global Const M_MODIFIED_ACCESS_RECTANGULAR_ON = &H1&
Global Const M_MODIFIED_ACCESS_OFF = 0
Global Const M_MODIFIED_ACCESS_ON = &H2&
Global Const M_MODIFIED_ACCESS_COMMAND_PROMPT = &H4&

'/* For hook after modification  */
Global Const M_HOOK_AFTER = &H10000000
'/* For hook before modification */
Global Const M_HOOK_BEFORE = &H20000000
'/* For buffer bitmap modification  */
Global Const M_HOOK_MODIFIED_DIB = 1&
'/* For disp window modification */
Global Const M_HOOK_MODIFIED_WINDOW = 2&
'/* For disp window modification */
Global Const M_HOOK_MESSAGE_LOOP = 4&
'/* For disp frame start */
Global Const M_FRAME_START = 9&
                                                        
'/************************************************************************/
'/* MdispOverlayKey()                                                    */
'/************************************************************************/
Global Const M_KEY_ON_COLOR = 1&
Global Const M_KEY_OFF = 2&
Global Const M_KEY_ALWAYS = 3&


'/************************************************************************/
'/* MdigAlloc() defines                                                  */
'/************************************************************************/
Global Const M_DIGITIZER_COLOR = &H1&
Global Const M_DIGITIZER_MONO = &H2&
Global Const M_USE_MEMORY_DCF = &H10&
Global Const M_DCF_REALLOC = &H20&

'/************************************************************************/
'/* MdigInquire() / MdigControl() Types                                  */
'/************************************************************************/

Global Const M_CHANNEL = 4000&
'/* Reserve next 1 bits                         from  (4000L | 0x00800000L)*/
Global Const M_CHANNEL_NUM = 4001&
Global Const M_BLACK_REF = 4003&
'/* Reserve next 8 bits                         from  (4003L | 0x00000000L)*/
'/*                                                   (4003L | 0x10000000L)*/
'/*                                                   (4003L | 0x20000000L)*/
'/*                                                   (4003L | 0x40000000L)*/
'/*                                                   (4003L | 0x80000000L)*/
'/*                                                   (4003L | 0x01000000L)*/
'/*                                                   (4003L | 0x02000000L)*/
'/*                                                   (4003L | 0x04000000L)*/
'/*                                             to    (4003L | 0x08000000L)*/
Global Const M_WHITE_REF = 4005&
'/* Reserve next 8 bits                         from  (4005L | 0x00000000L)*/
'/*                                                   (4005L | 0x10000000L)*/
'/*                                                   (4005L | 0x20000000L)*/
'/*                                                   (4005L | 0x40000000L)*/
'/*                                                   (4005L | 0x80000000L)*/
'/*                                                   (4005L | 0x01000000L)*/
'/*                                                   (4005L | 0x02000000L)*/
'/*                                                   (4005L | 0x04000000L)*/
'/*                                             to    (4005L | 0x08000000L)*/
Global Const M_HUE_REF = 4006&
Global Const M_SATURATION_REF = 4007&
Global Const M_BRIGHTNESS_REF = 4008&
Global Const M_CONTRAST_REF = 4009&
Global Const M_GRAB_SCALE = 4010&
Global Const M_GRAB_SCALE_X = 4011&
Global Const M_GRAB_SCALE_Y = 4012&
Global Const M_GRAB_SUBSAMPLE = 4013&
Global Const M_GRAB_SUBSAMPLE_X = 4014&
Global Const M_GRAB_SUBSAMPLE_Y = 4015&
Global Const M_GRAB_MODE = 4016&
Global Const M_GRAB_FRAME_NUM = 4017&
Global Const M_GRAB_FIELD_NUM = 4018&
Global Const M_GRAB_INPUT_GAIN = 4019&
Global Const M_INPUT_MODE = 4020&
Global Const M_SCAN_MODE = 4021&
Global Const M_SOURCE_SIZE_X = 4022&
Global Const M_SOURCE_SIZE_Y = 4023&
Global Const M_SOURCE_OFFSET_X = 4024&
Global Const M_SOURCE_OFFSET_Y = 4025&
Global Const M_INTERNAL_SOURCE_SIZE_X = 4026&
Global Const M_INTERNAL_SOURCE_SIZE_Y = 4027&
Global Const M_INTERNAL_SOURCE_OFFSET_X = 4028&
Global Const M_INTERNAL_SOURCE_OFFSET_Y = 4029&
Global Const M_GRAB_END_HANDLER_PTR = 4030&
Global Const M_GRAB_END_HANDLER_USER_PTR = 4032&
Global Const M_GRAB_START_HANDLER_PTR = 4033&
Global Const M_GRAB_START_HANDLER_USER_PTR = 4035&
Global Const M_GRAB_FIELD_END_HANDLER_PTR = 4036&
Global Const M_GRAB_FIELD_END_HANDLER_USER_PTR = 4037&
Global Const M_GRAB_FIELD_END_ODD_HANDLER_PTR = 4038&
Global Const M_GRAB_FIELD_END_ODD_HANDLER_USER_PTR = 4039&
Global Const M_GRAB_FIELD_END_EVEN_HANDLER_PTR = 4040&
Global Const M_GRAB_FIELD_END_EVEN_HANDLER_USER_PTR = 4041&
Global Const M_GRAB_FRAME_END_HANDLER_PTR = 4042&
Global Const M_GRAB_FRAME_END_HANDLER_USER_PTR = 4043&
Global Const M_GRAB_FRAME_START_HANDLER_PTR = 4044&
Global Const M_GRAB_FRAME_START_HANDLER_USER_PTR = 4045&
Global Const M_FIELD_START_HANDLER_PTR = 4046&
Global Const M_FIELD_START_HANDLER_USER_PTR = 4047&
Global Const M_FIELD_START_ODD_HANDLER_PTR = 4048&
Global Const M_FIELD_START_ODD_HANDLER_USER_PTR = 4049&
Global Const M_FIELD_START_EVEN_HANDLER_PTR = 4050&
Global Const M_FIELD_START_EVEN_HANDLER_USER_PTR = 4051&
Global Const M_SCALING_Y_AVAILABLE = 4052&
Global Const M_GRAB_TRIGGER_SOURCE = 4053&
Global Const M_GRAB_TRIGGER_MODE = 4054&
Global Const M_NATIVE_CAMERA_ID = 4060&
Global Const M_VCR_INPUT_TYPE = 4061&
Global Const M_CLIP_SRC_SUPPORTED = 4062&
Global Const M_CLIP_DST_SUPPORTED = 4063&
Global Const M_HOOK_FUNCTION_SUPPORTED = 4064&
Global Const M_GRAB_WINDOW_RANGE_SUPPORTED = 4065&
Global Const M_GRAB_SCALE_X_SUPPORTED = 4066&
Global Const M_GRAB_SCALE_Y_SUPPORTED = 4067&
Global Const M_GRAB_8_BITS_SUPPORTED = 4068&
Global Const M_GRAB_15_BITS_SUPPORTED = 4069&
Global Const M_GRAB_32_BITS_SUPPORTED = 4070&
Global Const M_GRAB_EXTRA_LINE = 4071&
Global Const M_GRAB_ABORT = 4072&
Global Const M_GRAB_DESTRUCTIVE_IN_PROGRESS = 4073&
Global Const M_GRAB_START_MODE = 4074&
Global Const M_GRAB_WINDOW_RANGE = 4075&
Global Const M_INPUT_SIGNAL_PRESENT = 4078&
Global Const M_INPUT_SIGNAL_SOURCE = 4079&
Global Const M_FIELD_START_THREAD_ID = 4080&
Global Const M_GRAB_FIELD_END_ODD_THREAD_ID = 4081&
Global Const M_GRAB_FIELD_END_EVEN_THREAD_ID = 4082&
Global Const M_FIELD_START_THREAD_HANDLE = 4083&
Global Const M_GRAB_FIELD_END_ODD_THREAD_HANDLE = 4084&
Global Const M_GRAB_FIELD_END_EVEN_THREAD_HANDLE = 4085&
Global Const M_FORMAT_UPDATE = 4086&
Global Const M_USER_BIT = 4087&
'/* Reserve next 31 values                      from   4087L*/
'/*                                             to     4118L*/
Global Const M_GRAB_FAIL_CHECK = 4120&
Global Const M_GRAB_FAIL_STATUS = 4121&
Global Const M_GRAB_FAIL_RETRY_NUMBER = 4122&
Global Const M_GRAB_ON_ONE_LINE = 4123&
Global Const M_GRAB_WRITE_FORMAT = 4124&
Global Const M_GRAB_LUT_PALETTE = 4125&
Global Const M_GRAB_HALT_ON_NEXT_FIELD = 4126&
Global Const M_GRAB_TIMEOUT = 4127&
Global Const M_GRAB_IN_PROGRESS = 4128&
Global Const M_FIELD_START_HOOK_WHEN_GRAB_ONLY = 4129&
Global Const M_SOUND_VOLUME_REF = 4130&
Global Const M_SOUND_VOLUME_RIGHT_REF = 4131&
Global Const M_SOUND_VOLUME_LEFT_REF = 4132&
Global Const M_SOUND_TYPE_REF = 4133&
Global Const M_SOUND_TYPE_STATUS = 4134&
Global Const M_SOUND_BASS_REF = 4135&
Global Const M_SOUND_TREBLE_REF = 4136&
Global Const M_EXTERNAL_CHROMINANCE = 4137&
Global Const M_TUNER_FREQUENCY = 4138&
Global Const M_TUNER_STANDARD = 4139&
Global Const M_CLOCK_NOT_ALWAYS_VALID = 4140&
Global Const M_GRAB_LINESCAN_MODE = 4141&
Global Const M_GRAB_PERIOD = 4142&
Global Const M_OVERRIDE_ROUTER = 4143&
Global Const M_GRAB_EXPOSURE = 4150&
'/* Reserve next 8 values                       from   4151L*/
'/*                                             to     4158L*/
Global Const M_GRAB_EXPOSURE_SOURCE = 4160&
'/* Reserve next 8 values                       from   4161L*/
'/*                                             to     4168L*/
Global Const M_GRAB_EXPOSURE_MODE = 4170&
'/* Reserve next 8 values                       from   4171L*/
'/*                                             to     4178L*/
Global Const M_GRAB_EXPOSURE_TIME = 4180&
'/* Reserve next 8 values                       from   4181L*/
'/*                                             to     4188L*/
Global Const M_GRAB_EXPOSURE_TIME_DELAY = 4190&
'/* Reserve next 8 values                       from   4191L*/
'/*                                             to     4198L*/
Global Const M_GRAB_TRIGGER = 4200&
'/* Reserve next 8 values                       from   4201L*/
'/*                                             to     4208L*/
Global Const M_GRAB_EXPOSURE_BYPASS = 4210&
Global Const M_DCF_REALLOC_HANDLER_PTR = 4211&
Global Const M_DCF_REALLOC_HANDLER_USER_PTR = 4212&
Global Const M_USER_IN_FORMAT = 4213&
Global Const M_USER_OUT_FORMAT = 4214&
Global Const M_GRAB_RESTRICTION_CHECK = 4215&
Global Const M_LAST_GRAB_BUFFER = 4216&
Global Const M_NATIVE_LAST_GRAB_OSB_ID = 4217&
Global Const M_SYNCHRONIZE_ON_STARTED = 4218&
Global Const M_GRAB_WAIT = 4219&
Global Const M_GRAB_FIELD_START_HANDLER_PTR = 4220&
Global Const M_GRAB_FIELD_START_HANDLER_USER_PTR = 4221&
Global Const M_GRAB_FIELD_START_ODD_HANDLER_PTR = 4222&
Global Const M_GRAB_FIELD_START_ODD_HANDLER_USER_PTR = 4223&
Global Const M_GRAB_FIELD_START_EVEN_HANDLER_PTR = 4224&
Global Const M_GRAB_FIELD_START_EVEN_HANDLER_USER_PTR = 4225&
Global Const M_GRAB_16_BITS_SUPPORTED = 4226&
Global Const M_GRAB_24_BITS_SUPPORTED = 4227&
'// Vidcap-only control.
Global Const M_VIDCAP_REF_DLG = 4228&
                                                       
'// !!! MAP FOR OLD DEFINES
Global Const M_DIG_TYPE = M_TYPE
Global Const M_DIG_NUMBER = M_NUMBER
Global Const M_DIG_FORMAT = M_FORMAT
Global Const M_DIG_INIT_FLAG = M_INIT_FLAG
Global Const M_DIG_CHANNEL_NUM = M_CHANNEL_NUM
Global Const M_DIG_LUT = M_LUT_ID
Global Const M_DIG_REF_BLACK = M_BLACK_REF
Global Const M_DIG_REF_WHITE = M_WHITE_REF
Global Const M_DIG_REF_HUE = M_HUE_REF
Global Const M_DIG_REF_SATURATION = M_SATURATION_REF
Global Const M_DIG_REF_BRIGHTNESS = M_BRIGHTNESS_REF
Global Const M_DIG_REF_CONTRAST = M_CONTRAST_REF
Global Const M_DIG_BLACK_REF = M_BLACK_REF
Global Const M_DIG_WHITE_REF = M_WHITE_REF
Global Const M_DIG_HUE_REF = M_HUE_REF
Global Const M_DIG_SATURATION_REF = M_SATURATION_REF
Global Const M_DIG_BRIGHTNESS_REF = M_BRIGHTNESS_REF
Global Const M_DIG_CONTRAST_REF = M_CONTRAST_REF
Global Const M_DIG_INPUT_MODE = M_INPUT_MODE
Global Const M_DIG_GRAB_SCALE = M_GRAB_SCALE
Global Const M_DIG_GRAB_SCALE_X = M_GRAB_SCALE_X
Global Const M_DIG_GRAB_SCALE_Y = M_GRAB_SCALE_Y
Global Const M_DIG_GRAB_SUBSAMPLE = M_GRAB_SUBSAMPLE
Global Const M_DIG_GRAB_SUBSAMPLE_X = M_GRAB_SUBSAMPLE_X
Global Const M_DIG_GRAB_SUBSAMPLE_Y = M_GRAB_SUBSAMPLE_Y
Global Const M_DIG_GRAB_MODE = M_GRAB_MODE
Global Const M_DIG_GRAB_FRAME_NUM = M_GRAB_FRAME_NUM
Global Const M_DIG_GRAB_FIELD_NUM = M_GRAB_FIELD_NUM
Global Const M_DIG_SOURCE_SIZE_X = M_SOURCE_SIZE_X
Global Const M_DIG_SOURCE_SIZE_Y = M_SOURCE_SIZE_Y
Global Const M_DIG_SOURCE_OFFSET_X = M_SOURCE_OFFSET_X
Global Const M_DIG_SOURCE_OFFSET_Y = M_SOURCE_OFFSET_Y
Global Const M_DIG_USER_BIT = M_USER_BIT
Global Const M_DIG_INPUT_SIGNAL_PRESENT = M_INPUT_SIGNAL_PRESENT
Global Const M_DIG_INPUT_SIGNAL_SOURCE = M_INPUT_SIGNAL_SOURCE
Global Const M_DIG_SOUND_VOLUME_REF = M_SOUND_VOLUME_REF
Global Const M_DIG_SOUND_VOLUME_RIGHT_REF = M_SOUND_VOLUME_RIGHT_REF
Global Const M_DIG_SOUND_VOLUME_LEFT_REF = M_SOUND_VOLUME_LEFT_REF
Global Const M_DIG_SOUND_TYPE_REF = M_SOUND_TYPE_REF
Global Const M_DIG_SOUND_BASS_REF = M_SOUND_BASS_REF
Global Const M_DIG_SOUND_TREBLE_REF = M_SOUND_TREBLE_REF
Global Const M_DIG_FORMAT_UPDATE = M_FORMAT_UPDATE
Global Const M_DIG_CLIP_SRC_SUPPORTED = M_CLIP_SRC_SUPPORTED
Global Const M_DIG_CLIP_DST_SUPPORTED = M_CLIP_DST_SUPPORTED
Global Const M_DIG_HOOK_FUNCTION_SUPPORTED = M_HOOK_FUNCTION_SUPPORTED
Global Const M_GRAB_INTERLACED_MODE = M_SCAN_MODE
Global Const M_GRAB_THREAD_PRIORITY = M_THREAD_PRIORITY
Global Const M_HOOK_PRIORITY = M_THREAD_PRIORITY
Global Const M_GRAB_WINDOWS_RANGE = M_GRAB_WINDOW_RANGE
Global Const M_GRAB_WINDOWS_RANGE_SUPPORTED = M_GRAB_WINDOW_RANGE_SUPPORTED

Global Const M_HARDWARE_PORT0 = 16&
Global Const M_SOFTWARE = 20&
Global Const M_VSYNC = 23&
Global Const M_HSYNC = 22&




Global Const M_HW_TRIGGER = M_HARDWARE_PORT0
Global Const M_SW_TRIGGER = M_SOFTWARE
Global Const M_VSYNC_TRIGGER = M_VSYNC
Global Const M_HSYNC_TRIGGER = M_HSYNC
Global Const M_DIG_CHANNEL = M_CHANNEL



'/* Inquire Values */
Global Const M_DIGITAL = 0&
Global Const M_ANALOG = 1&
Global Const M_INTERLACE = 0&
Global Const M_PROGRESSIVE = 1&
Global Const M_LINESCAN = 2&
Global Const M_MONOCHROME = 0&
Global Const M_COMPOSITE = 1&
Global Const M_ACTIVATE = 1&

Global Const M_PULSAR_XXX = 19&
Global Const M_PULSAR = 20&
Global Const M_PULSAR_WITH_RS422 = 21&
Global Const M_PULSAR_RS422_JIG = 22&

             


'// Corona board type
Global Const M_CORONA_XXX = 19&
Global Const M_CORONA = 20&
Global Const M_CORONA_LITE = 21&
Global Const M_CORONA_VIA = 22&
Global Const M_CORONA_RR = 23&
Global Const M_CORONA_DEV = 24&
Global Const M_CORONA_NO_DISP = 25&

'// Corona board type redefinition
Global Const M_DEVICE_CORONA_XXX = M_CORONA_XXX
Global Const M_DEVICE_CORONA = M_CORONA
Global Const M_DEVICE_CORONA_LITE = M_CORONA_LITE
Global Const M_DEVICE_CORONA_VIA = M_CORONA_VIA
Global Const M_DEVICE_CORONA_RR = M_CORONA_RR
Global Const M_DEVICE_CORONA_DEV = M_CORONA_DEV
Global Const M_DEVICE_CORONA_NO_DISP = M_CORONA_NO_DISP

Global Const M_METEOR = 20&
Global Const M_METEOR_TV = 21&
Global Const M_METEOR_RGB = 22&
Global Const M_METEOR_PRO = 23&
Global Const M_METEOR_TV_PRO = 24&
Global Const M_METEOR_RGB_PRO = 25&

Global Const M_GENESIS = 30&
Global Const M_GENESIS_PRO = 31&
Global Const M_GENESIS_LC = 32&


'/************************************************************************/
'/* MdigControl() / MdigInquire() Values                                 */
'/************************************************************************/

Global Const M_TIMER1 = 1&
Global Const M_TIMER2 = 2&
Global Const M_TIMER3 = 3&
Global Const M_TIMER4 = 4&
Global Const M_TIMER5 = 5&
Global Const M_TIMER6 = 6&
Global Const M_TIMER7 = 7&
Global Const M_TIMER8 = 8&

Global Const M_ARM_CONTINUOUS = 9&
Global Const M_ARM_MONOSHOT = 10&
Global Const M_ARM_RESET = 11&
Global Const M_EDGE_RISING = 12&
Global Const M_EDGE_FALLING = 13&
Global Const M_LEVEL_LOW = 14&
Global Const M_LEVEL_HIGH = 15&
Global Const M_HARDWARE_PORT1 = 17&
Global Const M_HARDWARE_PORT_CAMERA = 18&
Global Const M_START_EXPOSURE = 19&
Global Const M_USER_DEFINED = 21&

Global Const M_FILL_DESTINATION = -1#
Global Const M_FILL_DISPLAY = -2#
Global Const M_SYNCHRONOUS = 1&
Global Const M_ASYNCHRONOUS = 2&
Global Const M_ASYNCHRONOUS_QUEUED = 3&

Global Const M_LUT_PALETTE0 = 0&
Global Const M_LUT_PALETTE1 = 1&
Global Const M_LUT_PALETTE2 = 2&
Global Const M_LUT_PALETTE3 = 3&
Global Const M_LUT_PALETTE4 = 4&
Global Const M_LUT_PALETTE5 = 5&
Global Const M_LUT_PALETTE6 = 6&
Global Const M_LUT_PALETTE7 = 7&

Global Const M_GAIN0 = 10&
Global Const M_GAIN1 = 11&
Global Const M_GAIN2 = 12&
Global Const M_GAIN3 = 13&
Global Const M_GAIN4 = 14&

Global Const M_TTL = 1&
Global Const M_RS422 = 2&

Global Const M_FINAL_GRAB = -9998&


'/************************************************************************/
'/* MdigChannel()                                                        */
'/************************************************************************/
Global Const M_TUNER_CHANNEL = &H100000
'/* Reserve next 126 values (M_TUNER_CHANNEL +   0L)   0x00100000L*/
'/*                         (M_TUNER_CHANNEL + 126L)   0x0010007EL*/
Global Const M_TUNER_BAND = &H200000
'/* Reserve next 2   values (M_TUNER_BAND | M_REGULAR) 0x00220000L*/
'/*                         (M_TUNER_BAND | M_CABLE  ) 0x00220001L*/
Global Const M_CH0 = &H20000000
Global Const M_CH1 = &H40000000
Global Const M_CH2 = &H80000000
Global Const M_CH3 = &H1000000
Global Const M_CH4 = &H2000000
Global Const M_CH5 = &H4000000
Global Const M_CH6 = &H8000000
Global Const M_CH7 = &H200000
Global Const M_SYNC = &H400000
Global Const M_SIGNAL = &H800000
Global Const M_RGB = 8&
Global Const M_YC = 9&
Global Const M_RCA = M_CH0
Global Const M_ALL_CHANNEL = (M_CH0 Or M_CH1 Or M_CH2 Or M_CH3 Or M_CH4 Or M_CH5 Or M_CH6 Or M_CH7)
Global Const M_REGULAR = &H20000
Global Const M_CABLE = &H20001

'// !!! MAP FOR OLD DEFINES
Global Const M_DIG_TUNER_CHANNEL = M_TUNER_CHANNEL
Global Const M_DIG_TUNER_BAND = M_TUNER_BAND


'/************************************************************************/
'/* MdigReference()                                                      */
'/************************************************************************/
Global Const M_BLACK = 0&
Global Const M_WHITE = 1&
Global Const M_STEREO = 0&
Global Const M_MONO = 1&

'/* See the Inquire for the M_CHx values */
Global Const M_CH0_REF = M_CH0
Global Const M_CH1_REF = M_CH1
Global Const M_CH2_REF = M_CH2
Global Const M_CH3_REF = M_CH3
Global Const M_CH4_REF = M_CH4
Global Const M_CH5_REF = M_CH5
Global Const M_CH6_REF = M_CH6
Global Const M_CH7_REF = M_CH7
Global Const M_ALL_REF = (M_CH0_REF Or M_CH1_REF Or M_CH2_REF Or M_CH3_REF Or M_CH4_REF Or M_CH5_REF Or M_CH6_REF Or M_CH7_REF)
Global Const M_MIN_LEVEL = 0&
Global Const M_MAX_LEVEL = 255&


'/************************************************************************/
'/* MdigGrabWait()                                                       */
'/************************************************************************/
Global Const M_GRAB_NEXT_FRAME = 1&
Global Const M_GRAB_NEXT_FIELD = 2&
Global Const M_GRAB_END = 3&

                                                      
'/************************************************************************/
'/* MdigHookFunction()                                                   */
'/************************************************************************/
Global Const M_UNHOOK = &H4000000
Global Const M_GRAB_START = 4&
Global Const M_GRAB_FRAME_END = 5&
Global Const M_GRAB_FIELD_END_ODD = 6&
Global Const M_GRAB_FIELD_END_EVEN = 7&
Global Const M_GRAB_FIELD_END = 8&
Global Const M_FIELD_START = 10&
Global Const M_FIELD_START_ODD = 11&
Global Const M_FIELD_START_EVEN = 12&
Global Const M_GRAB_FRAME_START = 13&
Global Const M_GRAB_FIELD_START = 14&
Global Const M_GRAB_FIELD_START_ODD = 15&
Global Const M_GRAB_FIELD_START_EVEN = 16&



'/************************************************************************/
'/* MgenLutFunction()                                                    */
'/************************************************************************/
Global Const M_LOG = &H0&
Global Const M_EXP = &H1&
Global Const M_SIN = &H2&
Global Const M_COS = &H3&
Global Const M_TAN = &H4&
Global Const M_QUAD = &H5&

'/************************************************************************/
'/* MgenWarpParameter()                                                  */
'/************************************************************************/
'/* 8 bits reserved for number of fractional bits added to M_FIXED_POINT */
Global Const M_WARP_MATRIX = &H100000
Global Const M_WARP_POLYNOMIAL = &H200000
Global Const M_WARP_LUT = &H400000
Global Const M_WARP_4_CORNER = &H800000
'/* Optional controls */
Global Const M_FIXED_POINT = &H4000&
Global Const M_OVERSCAN_ENABLE = &H8000&

Global Const M_ID_OFFSET_OF_DEFAULT_KERNEL = &H100000

Global Const M_OVERSCAN_DISABLE = (M_ID_OFFSET_OF_DEFAULT_KERNEL / 2&)
Global Const M_VERY_FAST = &H10000
Global Const M_FAST = &H40000
'/* Transforms */
Global Const M_RESIZE = 1&
Global Const M_ROTATE = 2&
Global Const M_SHEAR_X = 3&
Global Const M_SHEAR_Y = 4&
Global Const M_TRANSLATE = 5&

'/************************************************************************/
'/* MimGetResult()                                                       */
'/************************************************************************/
Global Const M_VALUE = 0&
Global Const M_POSITION_X = &H3400&
Global Const M_POSITION_Y = &H4400&
Global Const M_NB_EVENT = 5&

'/************************************************************************/
'/* MimInquire()                                                         */
'/************************************************************************/
Global Const M_RESULT_SIZE = 0&
Global Const M_RESULT_TYPE = 1&

'/************************************************************************/
'/* MimFindExtreme()                                                     */
'/************************************************************************/
Global Const M_MAX_VALUE = 1&
Global Const M_MIN_VALUE = 2&

'/************************************************************************/
'/* MimArith()                                                           */
'/************************************************************************/
Global Const M_CONSTANT = &H8000&
Global Const M_ADD = &H0&
Global Const M_ADD_CONST = (M_ADD Or M_CONSTANT)
Global Const M_SUB = &H1&
Global Const M_SUB_CONST = (M_SUB Or M_CONSTANT)
Global Const M_NEG_SUB = &HA&
Global Const M_CONST_SUB = (M_NEG_SUB Or M_CONSTANT)
Global Const M_SUB_ABS = &H11&
Global Const M_MIN = &H2000000
Global Const M_MIN_CONST = (M_MIN Or M_CONSTANT)
Global Const M_MAX = &H4000000
Global Const M_MAX_CONST = (M_MAX Or M_CONSTANT)
Global Const M_OR = &H16&
Global Const M_OR_CONST = (M_OR Or M_CONSTANT)
Global Const M_AND = &H17&
Global Const M_AND_CONST = (M_AND Or M_CONSTANT)
Global Const M_XOR = &H18&
Global Const M_XOR_CONST = (M_XOR Or M_CONSTANT)
Global Const M_NOR = &H19&
Global Const M_NOR_CONST = (M_NOR Or M_CONSTANT)
Global Const M_NAND = &H1A&
Global Const M_NAND_CONST = (M_NAND Or M_CONSTANT)
Global Const M_XNOR = &H1B&
Global Const M_XNOR_CONST = (M_XNOR Or M_CONSTANT)
Global Const M_NOT = &H14&
Global Const M_NEG = &H23&
Global Const M_ABS = &HC&
Global Const M_PASS = &H2&
Global Const M_CONST_PASS = (M_PASS Or M_CONSTANT)
Global Const M_MULT = &H100&
Global Const M_MULT_CONST = (M_MULT Or M_CONSTANT)
Global Const M_DIV = &H101&
Global Const M_DIV_CONST = (M_DIV Or M_CONSTANT)
Global Const M_INV_DIV = &H102&
Global Const M_CONST_DIV = (M_INV_DIV Or M_CONSTANT)


'/************************************************************************/
'/* MimArithMultiple()                                                   */
'/************************************************************************/
Global Const M_OFFSET_GAIN = &H0&
Global Const M_WEIGHTED_AVERAGE = &H1&
Global Const M_MULTIPLY_ACCUMULATE_1 = &H2&
Global Const M_MULTIPLY_ACCUMULATE_2 = &H4&

'/************************************************************************/
'/* MimFlip()                                                            */
'/************************************************************************/
Global Const M_FLIP_VERTICAL = 1&
Global Const M_FLIP_HORIZONTAL = 2&

'/************************************************************************/
'/* MimBinarize(), MimClip(), MimLocateEvent()                           */
'/************************************************************************/
Global Const M_IN_RANGE = 1&
Global Const M_OUT_RANGE = 2&
Global Const M_EQUAL = 3&
Global Const M_NOT_EQUAL = 4&
Global Const M_GREATER = 5&
Global Const M_LESS = 6&
Global Const M_GREATER_OR_EQUAL = 7&
Global Const M_LESS_OR_EQUAL = 8&

'/************************************************************************/
'/* MimConvolve()                                                        */
'/************************************************************************/
Global Const M_SMOOTH = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 0&)
Global Const M_LAPLACIAN_EDGE = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 1&)
Global Const M_LAPLACIAN_EDGE2 = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 2&)
Global Const M_SHARPEN = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 3&)
Global Const M_SHARPEN2 = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 4&)
Global Const M_HORIZ_EDGE = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 5&)
Global Const M_VERT_EDGE = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 6&)
Global Const M_EDGE_DETECT = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 7&)
Global Const M_EDGE_DETECT2 = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 8&)

'/************************************************************************/
'/* MimEdgeDetect()                                                      */
'/************************************************************************/
Global Const M_SOBEL = M_EDGE_DETECT
Global Const M_PREWITT = M_EDGE_DETECT2
Global Const M_NOT_WRITE_ANGLE = 1&
Global Const M_NOT_WRITE_INT = 2&
Global Const M_FAST_ANGLE = 4&
Global Const M_FAST_GRADIENT = 8&
Global Const M_FAST_EDGE_DETECT = (M_FAST_ANGLE + M_FAST_GRADIENT)
Global Const M_REGULAR_ANGLE = 16&
Global Const M_REGULAR_GRADIENT = 64&
Global Const M_REGULAR_EDGE_DETECT = (M_REGULAR_ANGLE + M_REGULAR_GRADIENT)


'/************************************************************************/
'/* MimRank()                                                            */
'/************************************************************************/
Global Const M_MEDIAN = &H10000
Global Const M_3X3_RECT = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 20&)
Global Const M_3X3_CROSS = (M_ID_OFFSET_OF_DEFAULT_KERNEL + 21&)

'/************************************************************************/
'/* MimMorphic(), ...                                                    */
'/************************************************************************/
Global Const M_ERODE = 1&
Global Const M_DILATE = 2&
Global Const M_THIN = 3&
Global Const M_THICK = 4&
Global Const M_HIT_OR_MISS = 5&
Global Const M_MATCH = 6&
Global Const M_CLOSE = 10&
Global Const M_OPEN = 11&

'/************************************************************************/
'/* MimThin()                                                            */
'/************************************************************************/
Global Const M_TO_SKELETON = -1&


'/************************************************************************/
'/* MimThick()                                                           */
'/************************************************************************/
Global Const M_TO_IDEMPOTENCE = M_TO_SKELETON

'/************************************************************************/
'/* MimDistance()                                                        */
'/************************************************************************/
Global Const M_CHAMFER_3_4 = &H1
Global Const M_CITY_BLOCK = &H2
Global Const M_FORWARD = &H1
Global Const M_BACKWARD = &H2
Global Const M_OVERSCAN_TO_DO = &H4
Global Const M_BOTH = &H7

'/************************************************************************/
'/* MimProject()                                                         */
'/************************************************************************/
Global Const M_0_DEGREE = 0#
Global Const M_90_DEGREE = 90#
Global Const M_180_DEGREE = 180#
Global Const M_270_DEGREE = 270#

'/************************************************************************/
'/* MimResize(), MimTranslate() and MimRotate()                          */
'/************************************************************************/
Global Const M_INTERPOLATE = &H4&
Global Const M_BILINEAR = &H8&
Global Const M_BICUBIC = &H10&
Global Const M_AVERAGE = &H20&
Global Const M_NEAREST_NEIGHBOR = &H40&
Global Const M_OVERSCAN_CLEAR = &H80&
Global Const M_FIT_ALL_ANGLE = &H100&
Global Const M_BINARY = &H1000&

'/************************************************************************/
'/* MimAverage                                                           */
'/************************************************************************/
Global Const M_NORMAL = 1&
Global Const M_WEIGHTED = 2&
Global Const M_CONTINUOUS = -1&

'/************************************************************************/
'/* MimResize                                                            */
'/************************************************************************/


'/************************************************************************/
'/* MimHistogramEqualize()                                               */
'/************************************************************************/
Global Const M_UNIFORM = 1
Global Const M_EXPONENTIAL = 2
Global Const M_RAYLEIGH = 3
Global Const M_HYPER_CUBE_ROOT = 4
Global Const M_HYPER_LOG = 5


'/************************************************************************/
'/* MimConvert()                                                         */
'/************************************************************************/
Global Const M_RGB_TO_HLS = 1
Global Const M_RGB_TO_L = 2
Global Const M_HLS_TO_RGB = 3
Global Const M_L_TO_RGB = 4

'/************************************************************************/
'/* MimWarp()                                                            */
'/************************************************************************/
'/* 8 bits reserved for number of fractional bits */

'/************************************************************************/
'/* MimFFT()                                                             */
'/************************************************************************/
Global Const M_REVERSE = 4&
Global Const M_NORMALIZED = 2&
Global Const M_NORMALIZE = M_NORMALIZED
Global Const M_1D_FFT_ROWS = &H10&
Global Const M_1D_FFT_COLUMNS = &H20&


'/************************************************************************/
'/* Buffer attributes                                                    */
'/************************************************************************/
Global Const M_NO_ATTR = &H0
Global Const M_IN = &H1
Global Const M_OUT = &H2
Global Const M_SRC = M_IN
Global Const M_DEST = M_OUT

Global Const M_IMAGE = &H4&
Global Const M_GRAB = &H8&
Global Const M_PROC = &H10&
Global Const M_DISP = &H20&
Global Const M_GRAPH = &H40&
Global Const M_OVR = &H80&

Global Const M_FOR_SALE = &H4000&
Global Const M_MMX_ENABLED = &H8000&
Global Const M_FLIP = &H10000
Global Const M_PACKED = &H20000
Global Const M_PLANAR = &H40000
Global Const M_ON_BOARD = &H80000
Global Const M_OFF_BOARD = &H100000
Global Const M_NON_PAGED = &H200000
Global Const M_DIB = &H400000
Global Const M_SINGLE = &H1000000
Global Const M_VIA = M_SINGLE
Global Const M_PAGED = &H2000000
Global Const M_NO_FLIP = &H4000000
Global Const M_OVERSCAN_BUFFER = &H8000000


'/* 8 bits reserved for buffer internal format and format definitions */
Global Const M_INTERNAL_FORMAT = &H3F00&
Global Const M_INTERNAL_FORMAT_SHIFT = 8&
Global Const M_INTERNAL_COMPLETE_FORMAT = &HFFFFF00
Global Const M_ANY_INTERNAL_FORMAT = (0 * 2 ^ M_INTERNAL_FORMAT_SHIFT)
Global Const M_MONO1 = (1 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  1 bit  pixels in monochrome format
Global Const M_MONO8 = (2 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  8 bits pixels in monochrome format
Global Const M_MONO16 = (3 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 16 bits pixels in monochrome format
Global Const M_MONO32 = (4 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 32 bits pixels in monochrome format
Global Const M_RGB1 = (5 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                             '// Force  3 bits pixels in color RGB   1.1.1    format
Global Const M_RGB15 = (6 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force 16 bits pixels in color XRGB  1.5.5.5  format
Global Const M_RGB16 = (7 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force 16 bits pixels in color RGB   5.6.5    format
Global Const M_RGB24 = (8 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force 24 bits pixels in color RGB   8.8.8    format
Global Const M_RGB32 = (9 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force 32 bits pixels in color XRGB  8.8.8.8  format
Global Const M_RGB32_ATI = (10 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                       '// Force 32 bits pixels in color RGBX  8.8.8.8  format
Global Const M_RGB48 = (11 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 48 bits pixels in color RGB  16.16.16  format
Global Const M_RGB96 = (12 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 96 bits pixels in color RGB  32.32.32  format
Global Const M_BGR1 = (13 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  3 bits pixels in color RGB   1.1.1    format
Global Const M_BGR15 = (14 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 16 bits pixels in color BGRX  5.5.5.1  format
Global Const M_BGR16 = (15 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 16 bits pixels in color BGR   5.6.5    format
Global Const M_BGR24 = (16 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 24 bits pixels in color BGR   8.8.8    format
Global Const M_BGR32 = (17 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 32 bits pixels in color BGRX  8.8.8.8  format
Global Const M_BGR32_ATI = (18 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                       '// Force 32 bits pixels in color XBGR  8.8.8.8  format
Global Const M_BGR48 = (19 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 48 bits pixels in color BGR  16.16.16  format
Global Const M_BGR96 = (20 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 96 bits pixels in color BGR  32.32.32  format
Global Const M_YUV9 = (21 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  9 bits YUV pixels in color YUV 16:1:1 format
Global Const M_YUV12 = (22 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 12 bits YUV pixels in color YUV  4:1:1 format
Global Const M_YUV16 = (23 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 16 bits YUV pixels in color YUV  4:2:2 format
Global Const M_MONO8_VIA_RGB = (24 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                   '// Force  8 bits pixels in monochrome format from VIA RGB
Global Const M_RGB3 = (25 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  3 bits pixels in color RGB   1.1.1    format
Global Const M_BGR3 = (26 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                            '// Force  3 bits pixels in color BGR   1.1.1    format
Global Const M_YUV24 = (27 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                           '// Force 24 bits YUV pixels in color YUV  4:4:4 format
Global Const M_SINGLE_BAND = (255 * 2 ^ M_INTERNAL_FORMAT_SHIFT)                    '// PutColor and GetColor specification for a single band

'// !!! MAP FOR OLD DEFINES
Global Const M_CHAR = (M_MONO8 Or M_SIGNED)
Global Const M_UCHAR = (M_MONO8)
Global Const M_SHORT = (M_MONO16 Or M_SIGNED)
Global Const M_USHORT = (M_MONO16)
Global Const M_LONG = (M_MONO32 Or M_SIGNED)
Global Const M_ULONG = (M_MONO32)
Global Const M_RGB111 = (M_RGB1 Or M_SIGNED)
Global Const M_URGB111 = (M_RGB1)
Global Const M_BGR111 = (M_BGR1 Or M_SIGNED)
Global Const M_RGB555 = (M_RGB15 Or M_SIGNED)
Global Const M_URGB555 = (M_RGB15)
Global Const M_BGR555 = (M_BGR15 Or M_SIGNED)
Global Const M_RGB888 = (M_RGB24 Or M_SIGNED)
Global Const M_URGB888 = (M_RGB24)
Global Const M_BGR888 = (M_BGR24 Or M_SIGNED)
Global Const M_RGB161616 = (M_RGB48 Or M_SIGNED)
Global Const M_URGB161616 = (M_RGB48)
Global Const M_RGB323232 = (M_RGB96 Or M_SIGNED)
Global Const M_URGB323232 = (M_RGB96)
Global Const M_YUV9_PLANAR = (M_YUV9 Or M_PLANAR)
Global Const M_YUV12_PLANAR = (M_YUV12 Or M_PLANAR)
Global Const M_YUV16_PLANAR = (M_YUV16 Or M_PLANAR)
Global Const M_YUV16_PACKED = (M_YUV16 Or M_PACKED)
Global Const M_BGR15_PACKED = (M_BGR15 Or M_PACKED)
Global Const M_RGB15_PACKED = (M_RGB15 Or M_PACKED)
Global Const M_RGB24_PACKED = (M_RGB24 Or M_PACKED)
Global Const M_RGB32_PACKED = (M_RGB32 Or M_PACKED)
Global Const M_BGR24_PACKED = (M_BGR24 Or M_PACKED)
Global Const M_BGR16_PACKED = (M_BGR16 Or M_PACKED)
Global Const M_NODIBFLIP = (M_FLIP)
Global Const M_DIB_BGR16_PACKED = (M_BGR16 Or M_PACKED Or M_FLIP Or M_DIB)
Global Const M_DIB_BGR24_PACKED = (M_BGR24 Or M_PACKED Or M_FLIP Or M_DIB)
Global Const M_BGR32_PACKED = (M_BGR32 Or M_PACKED)
Global Const M_BGR32_PACKED_ATI = (M_BGR32_ATI Or M_PACKED)

Global Const M_LUT = &H100&
Global Const M_KERNEL = &H200&
Global Const M_STRUCT_ELEMENT = &H400&
Global Const M_ARRAY = &H800&
Global Const M_HIST_LIST = &H2000&
Global Const M_PROJ_LIST = &H4000&
Global Const M_EVENT_LIST = &H8000&
Global Const M_EXTREME_LIST = &H10000
Global Const M_COUNT_LIST = &H20000
Global Const M_FILE_FORMAT = &H40000
Global Const M_WARP_COEFFICIENT = &H80000
Global Const M_DIGITIZER = &H100000
Global Const M_DISPLAY = &H200000
Global Const M_APPLICATION = &H400000
Global Const M_SYSTEM = &H800000
Global Const M_GRAPHIC_CONTEXT = &H1000000
Global Const M_CALL_CONTEXT = &H2000000
Global Const M_ERROR_CONTEXT = &H4000000
Global Const M_USER_HOST_POINTER = &H10000000
Global Const M_USER_ATTRIBUTE = &H20000000
Global Const M_HOOK_CONTEXT = &H40000000

Global Const M_USER_OBJECT_1 = (M_USER_ATTRIBUTE Or &H10000)
Global Const M_USER_OBJECT_2 = (M_USER_ATTRIBUTE Or &H20000)
Global Const M_BLOB_OBJECT = (M_USER_ATTRIBUTE Or &H40000)
Global Const M_BLOB_FEATURE_LIST = (M_BLOB_OBJECT Or &H1&)
Global Const M_BLOB_RESULT = (M_BLOB_OBJECT Or &H2&)
Global Const M_PAT_OBJECT = (M_USER_ATTRIBUTE Or &H80000)
Global Const M_PAT_MODEL = (M_PAT_OBJECT Or &H1&)
Global Const M_PAT_RESULT = (M_PAT_OBJECT Or &H2&)
Global Const M_OCR_OBJECT = (M_USER_ATTRIBUTE Or &H100000)
Global Const M_OCR_FONT = (M_OCR_OBJECT Or &H1&)
Global Const M_OCR_RESULT = (M_OCR_OBJECT Or &H2&)
Global Const M_MEAS_OBJECT = (M_USER_ATTRIBUTE Or &H200000)
Global Const M_MEAS_MARKER = (M_MEAS_OBJECT Or &H1&)
Global Const M_MEAS_RESULT = (M_MEAS_OBJECT Or &H2&)
Global Const M_MEAS_CONTEXT = (M_MEAS_OBJECT Or &H4&)
Global Const M_MORE_MIL_OBJECT = (M_USER_ATTRIBUTE Or &H400000)
Global Const M_FREE_OBJECT_2 = (M_USER_ATTRIBUTE Or &H800000)
Global Const M_FREE_OBJECT_3 = (M_USER_ATTRIBUTE Or &H1000000)
Global Const M_FREE_OBJECT_4 = (M_USER_ATTRIBUTE Or &H2000000)
Global Const M_FREE_OBJECT_5 = (M_USER_ATTRIBUTE Or &H4000000)
Global Const M_FREE_OBJECT_6 = (M_USER_ATTRIBUTE Or &H8000000)
Global Const M_USER_DEFINE_LOW_ATTRIBUTE = &HFFFF&

Global Const M_THREAD_CONTEXT = (M_MORE_MIL_OBJECT Or &H1&)
Global Const M_EVENT_CONTEXT = (M_MORE_MIL_OBJECT Or &H2&)

Global Const M_SYSTEM_ALLOCATED = &H1&
Global Const M_USER_ALLOCATED = &HFFFFFFFE

'/************************************************************************/
'/* MbufCreateColor() Values                                             */
'/************************************************************************/
Global Const M_HOST_ADDRESS = &H80000000
Global Const M_PHYSICAL_ADDRESS = &H40000000
Global Const M_PITCH = &H20000000
Global Const M_PITCH_BYTE = &H8000000
Global Const M_BUF_ID = &H4000000
Global Const M_BUF_ID_MODIFY = &H2000000


'/************************************************************************/
'/* MbufGet(), MbufPut(), MbufChild(), ...                               */
'/************************************************************************/
Global Const M_RED = &H1000&
Global Const M_GREEN = &H2000&
Global Const M_BLUE = &H4000&
Global Const M_ALL_BAND = -1&
Global Const M_HUE = M_RED
Global Const M_SATURATION = M_GREEN
Global Const M_LUMINANCE = M_BLUE
Global Const M_Y = M_RED
Global Const M_U = M_GREEN
Global Const M_V = M_BLUE

Global Const M_ALL_BITS = -1&
Global Const M_DONT_CARE = &H8000&


'/************************************************************************/
'/* MbufImport(), MbufExport()                                           */
'/************************************************************************/
Global Const M_RESTORE = 0&
Global Const M_LOAD = 1&

Global Const M_MIL = 0&
Global Const M_RAW = 1&
Global Const M_TIFF = 2&
Global Const M_GIF = 3&


'/************************************************************************/
'/* MbufControlNeighborhood()                                            */
'/************************************************************************/
Global Const M_ABSOLUTE_VALUE = 50&
Global Const M_NORMALIZATION_FACTOR = 52&
Global Const M_OVERSCAN = 53&
Global Const M_OVERSCAN_REPLACE_VALUE = 54&
Global Const M_OFFSET_CENTER_X = 55&
Global Const M_OFFSET_CENTER_Y = 56&
Global Const M_TRANSPARENT = &H1000059
Global Const M_REPLACE = &H1000060
Global Const M_MIRROR = &H1000061
Global Const M_REPLACE_MAX = &H1000063
Global Const M_REPLACE_MIN = &H1000064


'/************************************************************************/
'/* MbufInquire() / MbufControl() Types                                  */
'/************************************************************************/

Global Const M_INTER_SYSTEM_ID = 5000&
Global Const M_PARENT_ID = 5001&
Global Const M_ANCESTOR_ID = 5002&
Global Const M_PARENT_OFFSET_X = 5003&
Global Const M_PARENT_OFFSET_Y = 5004&
Global Const M_ANCESTOR_OFFSET_X = 5005&
Global Const M_ANCESTOR_OFFSET_Y = 5006&
Global Const M_PARENT_OFFSET_BAND = 5007&
Global Const M_ANCESTOR_OFFSET_BAND = 5008&
Global Const M_NB_CHILD = 5009&
Global Const M_MODIFICATION_COUNT = 5010&
Global Const M_ANCESTOR_SIZE_X = 5012&
Global Const M_HOST_ADDRESS_FAR = 5013&
Global Const M_ASSOCIATED_LUT = 5014&
Global Const M_CURRENT_BUF_ID = 5015&
Global Const M_ASSOCIATED_BUFFER_ID = 5016&
Global Const M_MAP_BUFFER_TO_HOST = 5017&
Global Const M_HOST_ID = 5020&
Global Const M_DMA_BUFFER = 5021&
Global Const M_DMA_BUFFER_PTR = 5022&
Global Const M_DMA_BUFFER_PHYSICAL_PTR = 5023&
Global Const M_VALID_GRAB_BUFFER = 5025&
Global Const M_VALID_GRAB_BUFFER_OFFSET = 5026&
Global Const M_LOW_LEVEL_BUFFER_ID = 5027&
Global Const M_HOST_COLOR_ID = 5028&
Global Const M_MEMBANK = 5029&                                      '// Pulsar internal use only
Global Const M_LOCPOS_X = 5030&                                     '// Pulsar internal use only
Global Const M_LOCPOS_Y = 5031&                                     '// Pulsar internal use only
Global Const M_LOCPOS_BIT = 5032&                                   '// Pulsar internal use only
Global Const M_ON_BOARD_DISP_BUFFER_NATIVE_ID = 5033&               '// Pulsar internal use only
Global Const M_ON_BOARD_DISP_BUFFER_MIL_ID = 5034&                  '// Pulsar internal use only
Global Const M_VGA_DISP_BUFFER_ID = 5035&                           '// Pulsar internal use only
Global Const M_OVR_DISP_BUFFER_ID = 5036&                           '// Pulsar internal use only
Global Const M_MEMORG = 5037&                                       '// Pulsar internal use only
'// Free                                               5038L
Global Const M_DIB_MODE = 5039&
Global Const M_FLIP_MODE = 5040&
Global Const M_WINDOW_DC_ALLOC = 5041&
Global Const M_WINDOW_DC_FREE = 5042&
Global Const M_WINDOW_DC = 5043&
Global Const M_MODIFIED = 5044&

                                                        
'// !!! MAP FOR OLD DEFINES
Global Const M_DMA_READ_HOST_ID = M_HOST_ID
Global Const M_BUF_ASSOCIATED_BUFFER_ID = M_ASSOCIATED_BUFFER_ID
                                                        
'/************************************************************************/
'/* MbufControl() MbufInquire() Values                                   */
'/************************************************************************/

'/************************************************************************/
'/* MbufDiskInquire()                                                    */
'/************************************************************************/
Global Const M_LUT_PRESENT = 6000&
Global Const M_ASPECT_RATIO = 6001&


'/************************************************************************/
'/* Lattice values                                                       */
'/************************************************************************/
Global Const M_4_CONNECTED = &H10&
Global Const M_8_CONNECTED = &H20&


'/************************************************************************/
'/* Data types for results                                               */
'/************************************************************************/
Global Const M_TYPE_CHAR = &H10000
Global Const M_TYPE_SHORT = &H20000
Global Const M_TYPE_LONG = &H40000
Global Const M_TYPE_FLOAT = &H80000
Global Const M_TYPE_DOUBLE = &H100000
Global Const M_TYPE_PTR = &H200000
Global Const M_TYPE_MIL_ID = &H400000
Global Const M_TYPE_STRING = &H800000
Global Const M_TYPE_STRING_PTR = M_TYPE_STRING
Global Const M_TYPE_ASCII = M_TYPE_STRING
Global Const M_TYPE_BINARY = &H1000000
Global Const M_TYPE_HEX = &H2000000


'/* Bit encoded image types */
Global Const M_GREYSCALE = &H200&
Global Const M_GRAYSCALE = M_GREYSCALE

'/************************************************************************/
'/* MgraFont()                                                           */
'/************************************************************************/
Global Const M_FONT_DEFAULT_SMALL = 0&
Global Const M_FONT_DEFAULT_MEDIUM = 1&
Global Const M_FONT_DEFAULT_LARGE = 2&
Global Const M_FONT_DEFAULT = M_FONT_DEFAULT_SMALL
Global Const M_FONT_DEFAULT_SMALL_VALUE = 0&
Global Const M_FONT_DEFAULT_MEDIUM_VALUE = 1&
Global Const M_FONT_DEFAULT_LARGE_VALUE = 2&
Global Const M_FONT_DEFAULT_VALUE = M_FONT_DEFAULT_SMALL_VALUE

'/************************************************************************/
'/* MgraInquire()                                                        */
'/************************************************************************/
Global Const M_GRAPHIC_POSITION_X = 3&
Global Const M_GRAPHIC_POSITION_Y = 4&
Global Const M_COLOR = 5&
Global Const M_BACKCOLOR = 6&
Global Const M_FONT = 7&
Global Const M_FONT_X_SCALE = 8&
Global Const M_FONT_Y_SCALE = 9&
Global Const M_THICKNESS = 10&


'/************************************************************************/
'/* Used by MgraControl()                                                */
'/************************************************************************/
Global Const M_BACKGROUND_MODE = 12&
Global Const M_OPAQUE = &H1000058

'/************************************************************************/
'/* Used by MnatEnter/LeaveNativeMode()                                  */
'/************************************************************************/
Global Const M_NAT_NULL = &H0&
Global Const M_NAT_PROC = &H1&
Global Const M_NAT_GRAPH = &H2&
Global Const M_NAT_DISP = &H4&
Global Const M_NAT_GRAB = &H8&
Global Const M_NAT_ACCESS = &H10&
Global Const M_NAT_ALL = &H1F&

'/************************************************************************/
'/* Used by MnatAccessSystemInfo                                         */
'/************************************************************************/
Global Const M_MODULE_NAT = 0
Global Const M_READ = 1
Global Const M_WRITE = 2
Global Const M_RDWR = 3

'/************************************************************************/
'/* Used by MisNatGetLocGraph()                                          */
'/************************************************************************/
Global Const M_ORG = 1&
Global Const M_SURF = 2&

'/************************************************************************/
'/* MappGetError()                                                       */
'/************************************************************************/
Global Const M_NO_ERROR = 0&
Global Const M_CURRENT = 1&
Global Const M_CURRENT_FCT = 2&
Global Const M_CURRENT_SUB_NB = 3&
Global Const M_CURRENT_SUB = 4&
Global Const M_CURRENT_SUB_1 = 4&
Global Const M_CURRENT_SUB_2 = 5&
Global Const M_CURRENT_SUB_3 = 6&
Global Const M_GLOBAL = 7&
Global Const M_GLOBAL_FCT = 8&
Global Const M_GLOBAL_SUB_NB = 9&
Global Const M_GLOBAL_SUB = 10&
Global Const M_GLOBAL_SUB_1 = 10&
Global Const M_GLOBAL_SUB_2 = 11&
Global Const M_GLOBAL_SUB_3 = 12&
Global Const M_INTERNAL = 13&
Global Const M_INTERNAL_FCT = 14&
Global Const M_INTERNAL_SUB_NB = 15&
Global Const M_INTERNAL_SUB = 16&
Global Const M_INTERNAL_SUB_1 = 16&
Global Const M_INTERNAL_SUB_2 = 17&
Global Const M_INTERNAL_SUB_3 = 18&
Global Const M_PARAM_NB = 19&
Global Const M_FATAL = 40&
Global Const M_BUFFER_ID = 41&
Global Const M_REGION_OFFSET_X = 42&
Global Const M_REGION_OFFSET_Y = 43&
Global Const M_REGION_SIZE_X = 44&
Global Const M_REGION_SIZE_Y = 45&

Global Const M_MODIFIED_BUFFER = &H2000000
Global Const M_PARAM_VALUE = &H8000000
Global Const M_PARAM_TYPE = &H10000000
Global Const M_MESSAGE = &H20000000
Global Const M_ERROR = &H40000000
Global Const M_NATIVE_ERROR = &H80000000
Global Const M_THREAD_RECURSIVE = &H800000                                '/* Bit field exclusive to M_TRACE  to M_PROCESSING        */
Global Const M_THREAD_CURRENT = &H1000000                                 '/*                        M_TRACE_START                   */
'/*                        M_TRACE_END                     */
'/*                        M_ERROR                         */
'/*                        M_MESSAGE                       */
'/*                        M_CURRENT to M_REGION_SIZE_Y    */
'/*                        M_MODIFIED_BUFFER               */
'/*                        M_UNHOOK                        */
Global Const M_ERROR_CURRENT = (M_ERROR Or M_CURRENT)
Global Const M_ERROR_GLOBAL = (M_ERROR Or M_GLOBAL)
Global Const M_ERROR_FATAL = (M_ERROR Or M_FATAL)


'/************************************************************************/
'/* AppAlloc                                                             */
'/************************************************************************/
Global Const M_USER_OBJECT = &H1&
Global Const M_SYSTEM_OBJECT = &H2&
Global Const M_INHERITED = &H4&
Global Const M_NOT_INHERITED = &H8&
Global Const M_BROADCASTED = &H10&
Global Const M_NOT_BROADCASTED = &H20&

Global Const M_FUNCTION_NAME_SIZE = 32&
Global Const M_ERROR_FUNCTION_NAME_SIZE = M_FUNCTION_NAME_SIZE
Global Const M_ERROR_MESSAGE_SIZE = 128&

Global Const M_NBFCTNAMEMAX = 236                                 '/* max number of function codes     */
Global Const M_NBERRMSGMAX = 60                                   '/* max number of error messages     */
Global Const M_NBSUBERRMSGMAX = 10                                '/* max number of sub error messages */

Global Const M_FUNC_ERROR = (M_NBERRMSGMAX + 1&)                                 '/* M_MFUNC error numbers   */


'/************************************************************************/
'/* MappHookFunction()                                                   */
'/************************************************************************/
Global Const M_TRACE_START = 1&
Global Const M_TRACE_END = 2&
'/*                        M_TRACE_END                     */
'/*                        M_ERROR                         */
'/*                        M_MESSAGE                       */
'/*                        M_CURRENT to M_REGION_SIZE_Y    */
'/*                        M_MODIFIED_BUFFER               */
'/*                        M_UNHOOK                        */


'/************************************************************************/
'/* MappInquire() / MappControl() Types                                  */
'/************************************************************************/
Global Const M_VERSION = 1&
Global Const M_LAST_PLATFORM_USE = 7&
Global Const M_CURRENT_ERROR_HANDLER_PTR = 8&
Global Const M_CURRENT_ERROR_HANDLER_USER_PTR = 9&
Global Const M_GLOBAL_ERROR_HANDLER_PTR = 10&
Global Const M_GLOBAL_ERROR_HANDLER_USER_PTR = 11&
Global Const M_FATAL_ERROR_HANDLER_PTR = 12&
Global Const M_FATAL_ERROR_HANDLER_USER_PTR = 13&
Global Const M_TRACE_START_HANDLER_PTR = 14&
Global Const M_TRACE_START_HANDLER_USER_PTR = 15&
Global Const M_TRACE_END_HANDLER_PTR = 16&
Global Const M_TRACE_END_HANDLER_USER_PTR = 17&
Global Const M_IRQ_CONTROL = 18&
Global Const M_ERROR_HANDLER_PTR = 19&
Global Const M_ERROR_HANDLER_USER_PTR = 20&
Global Const M_MODIFIED_BUFFER_HANDLER_PTR = &H10000000                     '// Must not interfere with M_ERROR
Global Const M_MODIFIED_BUFFER_HANDLER_USER_PTR = &H20000000                '// Must not interfere with M_ERROR
Global Const M_OBJECT_TYPE = &H80000000


'/************************************************************************/
'/* MappInquire() / MappControl() Values                                 */
'/************************************************************************/
Global Const M_PARAMETER_CHECK = 1&
Global Const M_TRACE = 3&
Global Const M_PARAMETER = 4&
Global Const M_MEMORY = 5&
Global Const M_PROCESSING = 6&
Global Const M_PRINT_DISABLE = 0&
Global Const M_PRINT_ENABLE = 1&
Global Const M_CHECK_DISABLE = 2&
Global Const M_CHECK_ENABLE = 3&
Global Const M_COMPENSATION_DISABLE = 4&
Global Const M_COMPENSATION_ENABLE = 5&
'/*                        M_TRACE_END                     */
'/*                        M_ERROR                         */
'/*                        M_MESSAGE                       */
'/*                        M_CURRENT to M_REGION_SIZE_Y    */
'/*                        M_MODIFIED_BUFFER               */
'/*                        M_UNHOOK                        */

Global Const M_TIMER_ALLOC = 1&
Global Const M_TIMER_FREE = 2&
Global Const M_TIMER_RESOLUTION = 3&
Global Const M_TIMER_RESET = 4&
Global Const M_TIMER_READ = 5&
Global Const M_TIMER_WAIT = 6&
Global Const M_TIMER_MIL_NOP = &H8000&


'/************************************************************************/
'/* MappModify()                                                         */
'/************************************************************************/
Global Const M_SWAP_ID = 1&

'/************************************************************************/
'/* Binary functions in BLOB module.                                     */
'/************************************************************************/
Global Const M_LENGTH = &H2000&

'/************************************************************************/
'/* MmeasCalculate(), MmeasGetResult(), MpatGetResult() */
'/************************************************************************/
Global Const M_ANGLE = &H800&
Global Const M_ORIENTATION = &H2400&

'/************************************************************************/
'/* MblobControl() and/or MblobInquire() values and MmeasControl()       */
'/************************************************************************/
Global Const M_PIXEL_ASPECT_RATIO = 5&

'/************************************************************************/
'/* MfuncPrintMessage() defines                                          */
'/************************************************************************/
Global Const M_RESP_YES = 1&
Global Const M_RESP_NO = 2&
Global Const M_RESP_CANCEL = 4&
Global Const M_RESP_YES_NO = (M_RESP_YES Or M_RESP_NO)
Global Const M_RESP_YES_NO_CANCEL = (M_RESP_YES Or M_RESP_NO Or M_RESP_CANCEL)

'/************************************************************************/
'/* Mfile() defines                                                      */
'/************************************************************************/
Global Const M_NO_MEMORY = 1&


'/************************************************************************/
'* MsysAlloc
'/************************************************************************/
Global Const M_SYSTEM_HOST = M_SYSTEM_HOST_PTR
Global Const M_SYSTEM_MAGIC = M_SYSTEM_MAGIC_PTR
Global Const M_SYSTEM_IP8 = M_SYSTEM_IP8_PTR
Global Const M_SYSTEM_IMAGE = M_SYSTEM_IMAGE_PTR
Global Const M_SYSTEM_VGA = M_SYSTEM_VGA_PTR
Global Const M_SYSTEM_COMET = M_SYSTEM_COMET_PTR
Global Const M_SYSTEM_METEOR = M_SYSTEM_METEOR_PTR
Global Const M_SYSTEM_PULSAR = M_SYSTEM_PULSAR_PTR
Global Const M_SYSTEM_GENESIS = M_SYSTEM_GENESIS_PTR
Global Const M_SYSTEM_CORONA = M_SYSTEM_CORONA_PTR
Global Const M_SYSTEM_VIDCAP = M_SYSTEM_VIDCAP_PTR
Global Const M_SYSTEM_METEOR_II = M_SYSTEM_METEOR_II_PTR
'*************************************************************************




'******************************************************************************
'******************************************************************************
'******************************************************************************
'*
'* Filename:  MILBLOB.H
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   1.0  $
'* Content :  This file contains the defines for the MIL blob
'*            analysis module. (Mblob...).
'* COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
'* All Rights Reserved
'*
'*******************************************************************************
'*******************************************************************************
'*******************************************************************************


'/* Binary only */

Global Const M_LABEL_VALUE = 1&
Global Const M_AREA = 2&
Global Const M_PERIMETER = 3&
Global Const M_FERET_X = 4&
Global Const M_FERET_Y = 5&
Global Const M_BOX_X_MIN = 6&
Global Const M_BOX_Y_MIN = 7&
Global Const M_BOX_X_MAX = 8&
Global Const M_BOX_Y_MAX = 9&
Global Const M_FIRST_POINT_X = 10&
Global Const M_FIRST_POINT_Y = 11&
Global Const M_AXIS_PRINCIPAL_LENGTH = 12&
Global Const M_AXIS_SECONDARY_LENGTH = 13&
Global Const M_FERET_MIN_DIAMETER = 14&
Global Const M_FERET_MIN_ANGLE = 15&
Global Const M_FERET_MAX_DIAMETER = 16&
Global Const M_FERET_MAX_ANGLE = 17&
Global Const M_FERET_MEAN_DIAMETER = 18&
Global Const M_CONVEX_AREA = 19&
Global Const M_CONVEX_PERIMETER = 20&
Global Const M_X_MIN_AT_Y_MIN = 21&
Global Const M_X_MAX_AT_Y_MAX = 22&
Global Const M_Y_MIN_AT_X_MAX = 23&
Global Const M_Y_MAX_AT_X_MIN = 24&
Global Const M_COMPACTNESS = 25&
Global Const M_NUMBER_OF_HOLES = 26&
Global Const M_FERET_ELONGATION = 27&
Global Const M_ROUGHNESS = 28&
Global Const M_EULER_NUMBER = 47&
Global Const M_BREADTH = 49&
Global Const M_ELONGATION = 50&
Global Const M_INTERCEPT_0 = 51&
Global Const M_INTERCEPT_45 = 52&
Global Const M_INTERCEPT_90 = 53&
Global Const M_INTERCEPT_135 = 54&
Global Const M_NUMBER_OF_RUNS = 55&
Global Const M_GENERAL_FERET = &H400&

'/* Greyscale only (ie, trivial for binary) */

Global Const M_SUM_PIXEL = 29&
Global Const M_MIN_PIXEL = 30&
Global Const M_MAX_PIXEL = 31&
Global Const M_MEAN_PIXEL = 32&
Global Const M_SIGMA_PIXEL = 33&
Global Const M_SUM_PIXEL_SQUARED = 46&

'/* Binary or greyscale (might want both for a greyscale image) */

Global Const M_CENTER_OF_GRAVITY_X = 34&
Global Const M_CENTER_OF_GRAVITY_Y = 35&
Global Const M_MOMENT_X0_Y1 = 36&
Global Const M_MOMENT_X1_Y0 = 37&
Global Const M_MOMENT_X1_Y1 = 38&
Global Const M_MOMENT_X0_Y2 = 39&
Global Const M_MOMENT_X2_Y0 = 40&
Global Const M_MOMENT_CENTRAL_X1_Y1 = 41&
Global Const M_MOMENT_CENTRAL_X0_Y2 = 42&
Global Const M_MOMENT_CENTRAL_X2_Y0 = 43&
Global Const M_AXIS_PRINCIPAL_ANGLE = 44&
Global Const M_AXIS_SECONDARY_ANGLE = 45&
Global Const M_GENERAL_MOMENT = &H800&

'/* General moment type */

Global Const M_ORDINARY = &H400&
Global Const M_CENTRAL = &H800&

'/* Short cuts for enabling multiple features */

Global Const M_ALL_FEATURES = &H100&            '/* All except general Feret */
Global Const M_BOX = &H101&
Global Const M_CONTACT_POINTS = &H102&
Global Const M_CENTER_OF_GRAVITY = &H103&
Global Const M_NO_FEATURES = &H104&             '/* Still do label and area */

'/* MblobControl() and/or MblobInquire() values */

Global Const M_BLOB_IDENTIFICATION = 2&
Global Const M_LATTICE = 3&
Global Const M_FOREGROUND_VALUE = 4&
Global Const M_NUMBER_OF_FERETS = 6&
Global Const M_RESET = 9&
Global Const M_SAVE_RUNS = 14&
Global Const M_IDENTIFIER_TYPE = 15&
Global Const M_MAX_LABEL = 16&

'/* Blob identification values */

Global Const M_WHOLE_IMAGE = 1&
Global Const M_INDIVIDUAL = 2&
Global Const M_LABELLED = 4&

'/* Foreground values */

Global Const M_NONZERO = &H80&
Global Const M_ZERO = &H100&
Global Const M_NON_ZERO = M_NONZERO

'/* Conditional test not in MIL.H */

Global Const M_ALWAYS = 0&

'/* MblobReconstruct() defines */

Global Const M_RECONSTRUCT_FROM_SEED = 1&
Global Const M_ERASE_BORDER_BLOBS = 2&
Global Const M_FILL_HOLES = 3&
Global Const M_EXTRACT_HOLES = 4&
Global Const M_SEED_PIXELS_ALL_IN_BLOBS = 1&
Global Const M_FOREGROUND_ZERO = 2&

'/* Miscellaneous */

Global Const M_ALL_BLOBS = &H105&
Global Const M_INCLUDED_BLOBS = &H106&
Global Const M_EXCLUDED_BLOBS = &H107&
Global Const M_INCLUDE = 1&
Global Const M_EXCLUDE = 2&
Global Const M_DELETE = 3&
Global Const M_MIN_FERETS = 2&
Global Const M_MAX_FERETS = 64&
Global Const M_INCLUDE_ONLY = &H101&
Global Const M_EXCLUDE_ONLY = &H102&

'/* Other defines are in MIL.H */


'/********************************************************************
'* Error codes
'********************************************************************/

Global Const M_BLOB_RESULT_ID_ERROR = (2000 + M_FUNC_ERROR)
Global Const M_BLOB_FEATURE_ID_ERROR = (2001 + M_FUNC_ERROR)
                                      
'/* Memory errors */
Global Const M_BLOB_MEMORY_ERROR = (2050 + M_FUNC_ERROR)
Global Const M_BLOB_NO_IDS_ERROR = (2051 + M_FUNC_ERROR)

'/* Processing errors */
Global Const M_BLOB_PARAMETER_ERROR = (2100 + M_FUNC_ERROR)
Global Const M_BLOB_NOT_AVAILABLE_ERROR = (2101 + M_FUNC_ERROR)
Global Const M_BLOB_INTERNAL_ERROR = (2102 + M_FUNC_ERROR)
Global Const M_BLOB_TOO_COMPLEX_ERROR = (2103 + M_FUNC_ERROR)
Global Const M_BLOB_HW_LIMITATION_ERROR = (2104 + M_FUNC_ERROR)
Global Const M_BLOB_LIMITATION_ERROR = (2105 + M_FUNC_ERROR)


'/********************************************************************
'* Function prototypes
'********************************************************************/

Declare Function MblobAllocFeatureList Lib "milblob.dll" (ByVal MilSystem As Long, FeatureListPtr As Long) As Long
Declare Function MblobAllocResult Lib "milblob.dll" (ByVal MilSystem As Long, BlobResIdPtr As Long) As Long
Declare Function MblobGetLabel Lib "milblob.dll" (ByVal BlobResId As Long, ByVal XPos As Long, ByVal YPos As Long, BlobLabelPtr As Long) As Long
Declare Function MblobGetNumber Lib "milblob.dll" (ByVal BlobResId As Long, CountPtr As Long) As Long
Declare Sub MblobCalculate Lib "milblob.dll" (ByVal BlobIdentImageId As Long, ByVal GreyImageId As Long, ByVal FeatureListId As Long, ByVal BlobResId As Long)
Declare Sub MblobControl Lib "milblob.dll" (ByVal BlobResId As Long, ByVal ProcMode As Long, ByVal Value As Double)
Declare Sub MblobFill Lib "milblob.dll" (ByVal BlobResId As Long, ByVal TargetImageId As Long, ByVal Mode As Long, ByVal Value As Long)
Declare Sub MblobFree Lib "milblob.dll" (ByVal BlobResId As Long)
Declare Sub MblobGetResult Lib "milblob.dll" (ByVal BlobResId As Long, ByVal Feature As Long, TargetArrayPtr As Any)
Declare Sub MblobGetResultSingle Lib "milblob.dll" (ByVal BlobResId As Long, ByVal BlobLabel As Long, ByVal Feature As Long, ValuePtr As Any)
Declare Sub MblobGetRuns Lib "milblob.dll" (ByVal BlobResId As Long, ByVal BlobLabel As Long, ByVal ArrayType As Long, RunXPtr As Any, RunYPtr As Any, RunLengthPtr As Any)
Declare Sub MblobInquire Lib "milblob.dll" (ByVal BlobResId As Long, ByVal ProcMode As Long, ValuePtr As Any)
Declare Sub MblobLabel Lib "milblob.dll" (ByVal BlobResId As Long, ByVal TargetImageId As Long, ByVal Mode As Long)
Declare Sub MblobSelect Lib "milblob.dll" (ByVal BlobResId As Long, ByVal Operation As Long, ByVal Feature As Long, ByVal Condition As Long, ByVal CondLow As Double, ByVal CondHigh As Double)
Declare Sub MblobSelectFeature Lib "milblob.dll" (ByVal FeatureListId As Long, ByVal Feature As Long)
Declare Sub MblobSelectFeret Lib "milblob.dll" (ByVal FeatureListId As Long, ByVal Angle As Double)
Declare Sub MblobSelectMoment Lib "milblob.dll" (ByVal FeatureListId As Long, ByVal MomType As Long, ByVal XMomOrder As Long, ByVal YMomOrder As Long)
Declare Sub MblobReconstruct Lib "milblob.dll" (ByVal srce_image_id As Long, ByVal seed_image_id As Long, ByVal dest_image_id As Long, ByVal Operation As Long, ByVal Mode As Long)

'******************************************************************************




'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Filename:  MILMEAS.BAS
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   1.0  $
'* Content :  This file contains the defines for the MIL measurement
'*            module. (Mmeas...).
'* COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'******************************************************************************


'/**************************************************************************/
'/* CAPI defines                                                           */
'/**************************************************************************/

'/**************************************************************************/
'/* MmeasAllocMarker                                                       */
'/**************************************************************************/

Global Const M_POINT = 1&
Global Const M_EDGE = 2&
Global Const M_STRIPE = 3&


'/**************************************************************************/
'/* MmeasAllocResult                                                       */
'/**************************************************************************/
Global Const M_CALCULATE = 1&


'/**************************************************************************/
'/* Bitwise values that the followings cannot take                         */
'/**************************************************************************/

Global Const M_MULTI_MARKER_MASK = &H3FF&
Global Const M_EDGE_FIRST = &H100000
Global Const M_EDGE_SECOND = &H200000
Global Const M_WEIGHT_FACTOR = &H1000000
Global Const M_MEAN = &H3000000
Global Const M_MEAS_FUTURE_USE_0 = &H10000000
Global Const M_MEAS_FUTURE_USE_1 = &H20000000
Global Const M_MEAS_FUTURE_USE_2 = &H40000000
Global Const M_MEAS_FUTURE_USE_3 = &H80000000


'/**************************************************************************/
'/* MmeasInquire(), MmeasSetMarker(), MmeasGetResult(), MmeasFindMarker(), */
'/* MmeasCalculate() parameters :                                          */
'/**************************************************************************/

Global Const M_POSITION_VARIATION = &H8000&
Global Const M_WIDTH = &H10000
Global Const M_WIDTH_VARIATION = &H20000
Global Const M_POLARITY = &H4000&
Global Const M_CONTRAST = &H1000&
Global Const M_LINE_EQUATION = &H800000
Global Const M_LINE_EQUATION_SLOPE = &H801000
Global Const M_LINE_EQUATION_INTERCEPT = &H802000
Global Const M_EDGE_INSIDE = &H400000
Global Const M_POSITION = &H400&
Global Const M_SCORE = &H1400&
Global Const M_CONTRAST_VARIATION = &H5400&
Global Const M_EDGE_STRENGTH = &H6400&
Global Const M_EDGE_STRENGTH_VARIATION = &H7400&
Global Const M_EDGE_INSIDE_VARIATION = &H8400&
Global Const M_BOX_ORIGIN = &H9400&
Global Const M_BOX_SIZE = &HA400&
Global Const M_BOX_CENTER = &HB400&
Global Const M_BOX_FIRST_CENTER = &HC400&
Global Const M_BOX_SECOND_CENTER = &HD400&
Global Const M_BOX_ANGLE_MODE = &HE400&
Global Const M_BOX_ANGLE = &HF400&
Global Const M_BOX_ANGLE_DELTA_NEG = &H10400
Global Const M_BOX_ANGLE_DELTA_POS = &H11400
Global Const M_BOX_ANGLE_TOLERANCE = &H12400
Global Const M_BOX_ANGLE_ACCURACY = &H13400
Global Const M_BOX_ANGLE_INTERPOLATION_MODE = &H14400
Global Const M_EDGE_THRESHOLD = &H15400
Global Const M_MARKER_REFERENCE = &H16400
Global Const M_BOX_ANGLE_SIZE = &H17400
Global Const M_MARKER_TYPE = &H18400
Global Const M_CONTROL_FLAG = &H19400
Global Const M_POSITION_MIN = &H1A400
Global Const M_POSITION_MAX = &H1B400
Global Const M_BOX_EDGES_STRENGTH = &H1C400
Global Const M_ANY_ANGLE = &H1D400
Global Const M_VALID_FLAG = &H1E400
Global Const M_BOX_CORNER_TOP_LEFT = &H1F400
Global Const M_BOX_CORNER_TOP_RIGHT = &H20400
Global Const M_BOX_CORNER_BOTTOM_LEFT = &H21400
Global Const M_BOX_CORNER_BOTTOM_RIGHT = &H22400
Global Const M_BOX_EDGES_STRENGTH_NUMBER = &H23400
Global Const M_POSITION_INSIDE_STRIPE = &H24400
Global Const M_BOX_ANGLE_REFERENCE = &H25400

Global Const M_ZERO_OFFSET_X = 1&
Global Const M_ZERO_OFFSET_Y = 2&
Global Const M_PIXEL_ASPECT_RATIO_INPUT = 6&
Global Const M_PIXEL_ASPECT_RATIO_OUTPUT = 7&

Global Const M_DISTANCE = &H80000
Global Const M_DISTANCE_X = &H81000
Global Const M_DISTANCE_Y = &H82000


'/**************************************************************************/
'/* MmeasInquire(), MmeasSetMarker(), MmeasGetResult(), MmeasFindMarker(), */
'/* MmeasCalculate() values :                                              */
'/**************************************************************************/

Global Const M_VERTICAL = 1&
Global Const M_HORIZONTAL = 2&
Global Const M_ANY = &H11000000
Global Const M_POSITIVE = 2&
Global Const M_NEGATIVE = -2&          '/*Must be the additive inverse of M_POSITIVE*/
Global Const M_OPPOSITE = 3&
Global Const M_SAME = 4&
Global Const M_CORRECTED = 2&


'/**************************************************************************/
'/* Utility defines                                                        */
'/**************************************************************************/

Global Const M_INFINITE_SLOPE = (1E+300)


'/**************************************************************************/
'/* Function prototypes                                                    */
'/**************************************************************************/
Declare Function MmeasAllocMarker Lib "milmeas.dll" (ByVal SystemId As Long, ByVal MarkerType As Long, ByVal ControlFlag As Long, MarkerIdPtr As Long) As Long
Declare Function MmeasAllocResult Lib "milmeas.dll" (ByVal SystemId As Long, ByVal ResultBufferType As Long, ResultIdPtr As Long) As Long
Declare Function MmeasRestoreMarker Lib "milmeas.dll" (ByVal FileName As String, ByVal SystemId As Long, ByVal ControlFlag As Long, MarkerIdPtr As Long) As Long
Declare Function MmeasAllocContext Lib "milmeas.dll" (ByVal SystemId As Long, ByVal ControlFlag As Long, ContextId As Long) As Long
Declare Function MmeasInquire Lib "milmeas.dll" (ByVal MarkerIdOrResultIdOrContextId As Long, ByVal ParamToInquire As Long, FirstValuePtr As Any, SecondValuePtr As Any) As Long
Declare Sub MmeasFree Lib "milmeas.dll" (ByVal MarkerOrResultIdOrContextId As Long)
Declare Sub MmeasSaveMarker Lib "milmeas.dll" (ByVal FileName As String, ByVal MarkerId As Long, ByVal ControlFlag As Long)
Declare Sub MmeasSetMarker Lib "milmeas.dll" (ByVal MarkerId As Long, ByVal Parameter As Long, ByVal FirstValue As Double, ByVal SecondValue As Double)
Declare Sub MmeasFindMarker Lib "milmeas.dll" (ByVal ContextId As Long, ByVal ImageId As Long, ByVal MarkerId As Long, ByVal MeasurementList As Long)
Declare Sub MmeasCalculate Lib "milmeas.dll" (ByVal ContextId As Long, ByVal Marker1Id As Long, ByVal Marker2Id As Long, ByVal ResultId As Long, ByVal MeasurementList As Long)
Declare Sub MmeasGetResult Lib "milmeas.dll" (ByVal MarkerOrResultId As Long, ByVal ResultType As Long, FirstResultPtr As Any, SecondResultPtr As Any)
Declare Sub MmeasControl Lib "milmeas.dll" (ByVal ContextId As Long, ByVal ControlType As Long, ByVal Value As Double)

'******************************************************************************




'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Filename:  MILPAT.BAS
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   1.0  $
'* Content :  This file contains the defines for the MIL pattern
'*            recognition module. (Mpat...).
'* COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'******************************************************************************


'/* Bit encoded model types */
Global Const M_TEMPLATE = 1&
Global Const M_ROTATION = 4&
Global Const M_NOISY = &H800&

'/* Levels of speed and/or accuracy */
Global Const M_VERY_LOW = 0&
Global Const M_LOW = 1&
Global Const M_MEDIUM = 2&
Global Const M_HIGH = 3&
Global Const M_VERY_HIGH = 4&
Global Const M_FULL_SEARCH = &H80&

'/* Bit encoded flags for MpatPreprocModel() */
Global Const M_DELETE_LOW = &H10&
Global Const M_DELETE_MEDIUM = &H20&
Global Const M_DELETE_HIGH = &H40&

Global Const M_ALL = 0&
Global Const M_UNKNOWN = -9999&
Global Const M_NO_CHANGE = -9998&
Global Const M_ABSOLUTE = 1&
Global Const M_OFFSET = 2&

'/* 'type' parameter of MpatAlloc() */
Global Const M_NO_ROTATION = &H80000
Global Const M_BEST = &H100000
Global Const M_PAT_DEBUG = &H200000
Global Const M_MULTIPLE = &H400000

'/* 'Flag' parameter of MpatCopy() */
Global Const M_CLEAR_BACKGROUND = &H2000&

'/* Used by MpatGetResult() */
Global Const M_FOUND_FLAG = 1&
Global Const M_SCALE = 6&
Global Const M_ORIENTATION_SCORE = 7&

'/* Used by MpatInquire() */
Global Const M_ALLOC_TYPE = 1&
Global Const M_ALLOC_SIZE_X = 2&
Global Const M_ALLOC_SIZE_Y = 3&
Global Const M_CENTER_X = 4&
Global Const M_CENTER_Y = 5&
Global Const M_ORIGINAL_X = 6&
Global Const M_ORIGINAL_Y = 7&
Global Const M_SPEED_FACTOR = 8&
Global Const M_POSITION_START_X = 9&
Global Const M_POSITION_START_Y = 10&
Global Const M_POSITION_UNCERTAINTY_X = 11&
Global Const M_POSITION_UNCERTAINTY_Y = 12&
Global Const M_POSITION_ACCURACY = 13&
Global Const M_PREPROCESSED = 14&
Global Const M_ALLOC_OFFSET_X = 15&
Global Const M_ALLOC_OFFSET_Y = 16&
Global Const M_ACCEPTANCE_THRESHOLD = 17&
Global Const M_NUMBER_OF_OCCURENCES = 18&
Global Const M_WORKSPACE_SIZE_X = 19&
Global Const M_WORKSPACE_SIZE_Y = 20&
Global Const M_WORKSPACE_FAST = 21&
Global Const M_WORKSPACE_ROTATION = 22&
Global Const M_WORKSPACE = 23&
Global Const M_NUMBER_OF_ENTRIES = 24&
Global Const M_CERTAINTY_THRESHOLD = 25&

'/* Search parameters */
Global Const M_FIRST_LEVEL = 31&
Global Const M_LAST_LEVEL = 32&
Global Const M_MODEL_STEP = 33&
Global Const M_FAST_FIND = 34&
Global Const M_MIN_SPACING_X = 35&
Global Const M_MIN_SPACING_Y = 36&
Global Const M_SCORE_TYPE = 37&
Global Const M_MODEL_FILE_TYPE = 38&
Global Const M_MODEL_FILE_TYPE_HOST = 3&
Global Const M_MODEL_FILE_TYPE_GENESIS = 4&

'/* Parameters for find orientation */
Global Const M_RESULT_RANGE_180 = &H1&
Global Const M_RESULT_RANGE_90 = &H2&
Global Const M_RESULT_RANGE_360 = &H4&
Global Const M_RESULT_RANGE_45 = &H8&
Global Const M_ORIENTATION_ACCEPTANCE = 200#


'/* Search parameters for search with rotation */
Global Const M_SEARCH_ANGLE_MODE = &H80&
Global Const M_SEARCH_ANGLE = &H100&
Global Const M_SEARCH_ANGLE_DELTA_NEG = &H200&
Global Const M_SEARCH_ANGLE_DELTA_POS = &H400&
Global Const M_SEARCH_ANGLE_TOLERANCE = &H800&
Global Const M_SEARCH_ANGLE_ACCURACY = &H1000&
Global Const M_SEARCH_ANGLE_FINE_REGION = &H2000&
Global Const M_SEARCH_ANGLE_DEBUG = &H4000&
Global Const M_SEARCH_ANGLE_INTERPOLATION_MODE = &H8000&
Global Const M_SEARCH_ANGLE_DIRTY = &H10000
Global Const M_SEARCH_ANGLE_MAGIC_V1 = &H100CAFE
Global Const M_SEARCH_ANGLE_V1_FREE = &H128

Global Const M_DEF_SEARCH_ANGLE_MODE = M_DISABLE
Global Const M_DEF_SEARCH_ANGLE = 0#
Global Const M_DEF_SEARCH_ANGLE_DELTA_NEG = 0#
Global Const M_DEF_SEARCH_ANGLE_DELTA_POS = 0#
Global Const M_DEF_SEARCH_ANGLE_TOLERANCE = 5#
Global Const M_DEF_SEARCH_ANGLE_ACCURACY = M_DISABLE
Global Const M_DEF_SEARCH_ANGLE_FINE_REGION = 30&
Global Const M_DEF_SEARCH_ANGLE_DEBUG = M_DISABLE
Global Const M_DEF_SEARCH_ANGLE_INTERPOLATION_MODE = M_NEAREST_NEIGHBOR

'/* Spelling variations */
Global Const M_CENTRE_X = M_CENTER_X
Global Const M_CENTRE_Y = M_CENTER_Y

'/********************************************************************
'* Error codes
'********************************************************************/

'/* Model errors */
Global Const M_PAT_MODEL_ID_ERROR = (1000 + M_FUNC_ERROR)
Global Const M_PAT_MODEL_CORRUPT_ERROR = (1001 + M_FUNC_ERROR)
Global Const M_PAT_MODEL_TYPE_ERROR = (1002 + M_FUNC_ERROR)

'/* Result errors */
Global Const M_PAT_RESULT_ID_ERROR = (1050 + M_FUNC_ERROR)
Global Const M_PAT_RESULT_EMPTY_ERROR = (1051 + M_FUNC_ERROR)

'/* File errors */
Global Const M_PAT_FILE_OPEN_ERROR = (1100 + M_FUNC_ERROR)
Global Const M_PAT_FILE_READ_ERROR = (1101 + M_FUNC_ERROR)
Global Const M_PAT_FILE_WRITE_ERROR = (1102 + M_FUNC_ERROR)

'/* Memory errors */
Global Const M_PAT_MEMORY_ERROR = (1150 + M_FUNC_ERROR)
Global Const M_PAT_NO_IDS_ERROR = (1152 + M_FUNC_ERROR)

'/* Processing errors */
Global Const M_PAT_PARAMETER_ERROR = (1204 + M_FUNC_ERROR)
Global Const M_PAT_WORKBUF_ERROR = (1205 + M_FUNC_ERROR)
Global Const M_PAT_NOT_IMPLEMENTED_ERROR = (1207 + M_FUNC_ERROR)

'/********************************************************************
'* Function prototypes
'********************************************************************/
Declare Function MpatAlloc Lib "milpat.dll" (ByVal MilSystem As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal Type1 As Long, IdPtr As Long) As Long
Declare Function MpatAllocAutoModel Lib "milpat.dll" (ByVal MilSystem As Long, ByVal SrcImageId As Long, ByVal OffXPtr As Long, ByVal OffYPtr As Long, ByVal SizeXPtr As Long, ByVal SizeYPtr As Long, ByVal ModelType As Long, ByVal Mode As Long, IdPtr As Long) As Long
Declare Function MpatAllocModel Lib "milpat.dll" (ByVal MilSystem As Long, ByVal SrcImageId As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal ModelType As Long, IdPtr As Long) As Long
Declare Function MpatAllocResult Lib "milpat.dll" (ByVal MilSystem As Long, ByVal NumEntries As Long, IdPtr As Long) As Long
Declare Function MpatAllocRotatedModel Lib "milpat.dll" (ByVal MilSystem As Long, ByVal SrcModelorImageId As Long, ByVal Angle As Double, ByVal InterpolMode As Long, ByVal ModelType As Long, IdPtr As Long) As Long
Declare Function MpatGetNumber Lib "milpat.dll" (ByVal ResultId As Long, CountPtr As Long) As Long
Declare Function MpatRead Lib "milpat.dll" (ByVal MilSystem As Long, ByVal FileHandle As Integer, IdPtr As Long) As Long
Declare Function MpatRestore Lib "milpat.dll" (ByVal MilSystem As Long, ByVal FileName As String, IdPtr As Long) As Long
Declare Function MpatAllocRotatedModelAndSetDontCare Lib "milpat.dll" (ByVal MilSystem As Long, ByVal SrcImageId As Long, ByVal SrcDontCareId As Long, ByVal Angle As Double, ByVal InterpolMode As Long, ByVal ModelType As Long, IdPtr As Long) As Long
Declare Sub MpatCopy Lib "milpat.dll" (ByVal ModelId As Long, ByVal ImageId As Long, ByVal Version As Long)
Declare Sub MpatFindModel Lib "milpat.dll" (ByVal ImageId As Long, ByVal ModelId As Long, ByVal ResultId As Long)
Declare Sub MpatFindMultipleModel Lib "milpat.dll" (ByVal ImageId As Long, ModelId As Long, ResultId As Long, ByVal NumModels As Long, ByVal Flag As Long)
Declare Sub MpatFindOrientation Lib "milpat.dll" (ByVal ImageId As Long, ByVal ModelId As Long, ByVal ResultId As Long, ByVal ResultRange As Long)
Declare Sub MpatFree Lib "milpat.dll" (ByVal PatBufferId As Long)
Declare Sub MpatGetResult Lib "milpat.dll" (ByVal ResultId As Long, ByVal Type1 As Long, ArrayPtr As Double)
Declare Sub MpatInquire Lib "milpat.dll" (ByVal ModelId As Long, ByVal Item As Long, VarPtr As Any)
Declare Sub MpatPreprocModel Lib "milpat.dll" (ByVal ImageId As Long, ByVal ModelId As Long, ByVal Mode As Long)
Declare Sub MpatSave Lib "milpat.dll" (ByVal FileName As String, ByVal ModelId As Long)
Declare Sub MpatSetAcceptance Lib "milpat.dll" (ByVal ModelId As Long, ByVal AcceptanceThreshold As Double)
Declare Sub MpatSetAccuracy Lib "milpat.dll" (ByVal ModelId As Long, ByVal Accuracy As Long)
Declare Sub MpatSetAngle Lib "milpat.dll" (ByVal ModelId As Long, ByVal ControlType As Long, ByVal ControlValue As Double)
Declare Sub MpatSetCenter Lib "milpat.dll" (ByVal ModelId As Long, ByVal OffX As Double, ByVal OffY As Double)
Declare Sub MpatSetCertainty Lib "milpat.dll" (ByVal ModelId As Long, ByVal CertaintyThreshold As Double)
Declare Sub MpatSetDontCare Lib "milpat.dll" (ByVal ModelId As Long, ByVal ImageId As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal Value As Long)
Declare Sub MpatSetNumber Lib "milpat.dll" (ByVal ModelId As Long, ByVal NumMatches As Long)
Declare Sub MpatSetPosition Lib "milpat.dll" (ByVal ModelId As Long, ByVal StartX As Long, ByVal StartY As Long, ByVal SizeX As Long, ByVal SizeY As Long)
Declare Sub MpatSetSearchParameter Lib "milpat.dll" (ByVal ModelId As Long, ByVal Parameter As Long, ByVal Value As Double)
Declare Sub MpatSetSpeed Lib "milpat.dll" (ByVal ModelId As Long, ByVal Speed As Long)
Declare Sub MpatWrite Lib "milpat.dll" (FileHandle As Integer, ByVal ModelId As Long)

'******************************************************************************



'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Filename:  MILOCR.BAS
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   1.0  $
'* Rev     :  $Revision:   1.0  $
'*            Imaging Library (MIL) C OCR module user's functions.
'* COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'******************************************************************************

Declare Function MocrAllocFont Lib "milocr.dll" (ByVal SystemId As Long, ByVal FontType As Long, ByVal CharNumber As Long, ByVal CharBoxSizeX As Long, ByVal CharBoxSizeY As Long, ByVal CharOffsetX As Long, ByVal CharOffsetY As Long, ByVal CharSizeX As Long, ByVal CharSizeY As Long, ByVal CharThickness As Long, ByVal StringLength As Long, ByVal InitFlag As Long, FontIdPtr As Long) As Long
Declare Function MocrAllocResult Lib "milocr.dll" (ByVal SystemId As Long, ByVal InitFlag As Long, ResultIdPtr As Long) As Long
Declare Function MocrRestoreFont Lib "milocr.dll" (ByVal FileName As String, ByVal Operation As Long, ByVal SystemId As Long, FontId As Long) As Long
Declare Function MocrValidateString Lib "milocr.dll" (ByVal FontId As Long, ByVal Mode As Long, ByVal String1 As String) As Long
Declare Function OcrSemiCheckValid Lib "milocr.dll" (ByVal HookType As Long, ByVal String1 As String, FExpansionFlagPtr As Any) As Long
Declare Function OcrSemiM1292CheckValid Lib "milocr.dll" (ByVal HookType As Long, ByVal String1 As String, FExpansionFlagPtr As Any) As Long
Declare Function OcrSemiM1388CheckValid Lib "milocr.dll" (ByVal HookType As Long, ByVal String1 As String, FExpansionFlagPtr As Any) As Long
Declare Function OcrDefaultCheckValid Lib "milocr.dll" (ByVal HookType As Long, ByVal String1 As String, FExpansionFlagPtr As Any) As Long
Declare Sub MocrCalibrateFont Lib "milocr.dll" (ByVal CalibrationImageId As Long, ByVal FontId As Long, ByVal String1 As String, ByVal TargetCharSizeXMin As Double, ByVal TargetCharSizeXMax As Double, ByVal TargetCharSizeXStep As Double, ByVal TargetCharSizeYMin As Double, ByVal TargetCharSizeYmax As Double, ByVal TargetCharSizeYStep As Double, ByVal Operation As Long)
Declare Sub MocrCopyFont Lib "milocr.dll" (ByVal ImageId As Long, ByVal FontId As Long, ByVal Operation As Long, ByVal CharListString As String)
Declare Sub MocrFree Lib "milocr.dll" (ByVal FontOrResultId As Long)
Declare Sub MocrGetResult Lib "milocr.dll" (ByVal ResultId As Long, ByVal ResultToGet As Long, ResultPtr As Any)
Declare Sub MocrGetResultString Lib "milocr.dll" Alias "MocrGetResult" (ByVal ResultId As Long, ByVal ResultToGet As Long, ByVal ResultPtr As String)
Declare Sub MocrImportFont Lib "milocr.dll" (ByVal FileName As String, ByVal FileFormat As Long, ByVal Operation As Long, ByVal CharListString As String, ByVal FontId As Long)
Declare Sub MocrInquire Lib "milocr.dll" (ByVal FontId As Long, ByVal InquireItem As Long, ResultPtr As Any)
Declare Sub MocrInquireString Lib "milocr.dll" Alias "MocrInquire" (ByVal FontId As Long, ByVal InquireItem As Long, ByVal ResultPtr As String)
Declare Sub MocrReadString Lib "milocr.dll" (ByVal ImageId As Long, ByVal FontId As Long, ByVal ResultId As Long)
Declare Sub MocrModifyFont Lib "milocr.dll" (ByVal FontId As Long, ByVal Operation As Long, ByVal ControlFlag As Long)
Declare Sub MocrSaveFont Lib "milocr.dll" (ByVal FileName As String, ByVal Operation As Long, ByVal FontId As Long)
Declare Sub MocrSetConstraint Lib "milocr.dll" (ByVal FontId As Long, ByVal CharPos As Long, ByVal CharPosType As Long, ByVal CharValidString As String)
Declare Sub MocrControl Lib "milocr.dll" (ByVal FontId As Long, ByVal ControlToSet As Long, ByVal Value As Double)
Declare Sub MocrVerifyString Lib "milocr.dll" (ByVal ImageId As Long, ByVal FontId As Long, ByVal String1 As String, ByVal ResultId As Long)


'// Blob control defaults
Global Const M_DEF_STRING_LOC_MAX_NB_ITER = 2&
Global Const M_DEF_STRING_LOC_STOP_ITER = 0.5
Global Const M_DEF_STRING_LOC_GOOD_NB_CHAR = 4&
Global Const M_DEF_STRING_READ_BAD_SIZE_X = 0.4
Global Const M_DEF_STRING_READ_BAD_SIZE_Y = 0.4
Global Const M_DEF_STRING_READ_GOOD_SIZE_X = 0.25
Global Const M_DEF_STRING_READ_GOOD_SIZE_Y = 0.25
Global Const M_DEF_STRING_READ_BAD_ADD_CHAR = 4&
Global Const M_DEF_STRING_LOC_MIN_CHAR_SIZE = 0.66
Global Const M_DEF_STRING_LOC_MAX_CHAR_SIZE = 1.5
Global Const M_DEF_STRING_LOC_MIN_CHAR_SPACE = 0.66
Global Const M_DEF_STRING_LOC_MAX_CHAR_DISTANCE = 0.5
Global Const M_DEF_STRING_LOC_GOOD_CHAR_SIZE = 0.9
Global Const M_DEF_STRING_MAX_SLOPE = 0.1763269
Global Const M_DEF_STRING_READ_SIZE_X = 0.33
Global Const M_DEF_STRING_READ_SIZE_Y = 0.25
Global Const M_DEF_SKIP_SEARCH = M_DISABLE
Global Const M_DEF_SKIP_STRING_LOCATION = M_DISABLE
Global Const M_DEF_SKIP_CONTRAST_ENHANCE = M_DISABLE
Global Const M_DEF_STRING_ACCEPTANCE = 1#
Global Const M_DEF_CHAR_ACCEPTANCE = 1#
Global Const M_DEF_CHAR_INVALID = 0
Global Const M_DEF_DEBUG = 0
Global Const M_DEF_ADD_TOP_HAT = 2&
Global Const M_DEF_KILL_BORDER = M_ENABLE
Global Const M_DEF_READ_SPEED = M_MEDIUM
Global Const M_DEF_READ_ACCURACY = M_MEDIUM
Global Const M_DEF_READ_FIRST_LEVEL = M_DEFAULT
Global Const M_DEF_READ_LAST_LEVEL = M_DEFAULT
Global Const M_DEF_READ_MODEL_STEP = M_DEFAULT
Global Const M_DEF_READ_FAST_FIND = M_DEFAULT
Global Const M_DEF_READ_ROBUSTNESS = M_MEDIUM
Global Const M_DEF_STRING_LOC_NB_MODELS = 2&
Global Const M_DEF_PAT_ON_ACCELERATED = M_DISABLE
Global Const M_DEF_BLOB_ON_ACCELERATED = M_DISABLE
Global Const M_DEF_PROC_ON_ACCELERATED = M_ENABLE

'// Control associated InfoType defines
Global Const M_STRING_LOC_CHAR_SIZE_X = 1&
Global Const M_STRING_LOC_CHAR_SIZE_Y = 2&
Global Const M_STRING_LOC_MAX_NB_ITER = 3&
Global Const M_STRING_LOC_STOP_ITER = 5&
Global Const M_STRING_LOC_GOOD_NB_CHAR = 6&
Global Const M_STRING_READ_BAD_SIZE_X = 7&
Global Const M_STRING_READ_BAD_SIZE_Y = 8&
Global Const M_STRING_READ_GOOD_SIZE_X = 9&
Global Const M_STRING_READ_GOOD_SIZE_Y = 10&
Global Const M_STRING_READ_BAD_ADD_CHAR = 11&
Global Const M_STRING_LOC_MIN_CHAR_SIZE = 12&
Global Const M_STRING_LOC_MAX_CHAR_SIZE = 13&
Global Const M_STRING_LOC_MIN_CHAR_SPACE = 14&
Global Const M_STRING_LOC_MAX_CHAR_DISTANCE = 15&
Global Const M_STRING_LOC_GOOD_CHAR_SIZE = 16&
Global Const M_STRING_MAX_SLOPE = 17&
Global Const M_STRING_READ_SIZE_X = 18&
Global Const M_STRING_READ_SIZE_Y = 19&
Global Const M_SKIP_SEARCH = 21&
Global Const M_SKIP_STRING_LOCATION = 22&
Global Const M_SKIP_CONTRAST_ENHANCE = 23&
Global Const M_STRING_ACCEPTANCE = 24&
Global Const M_CHAR_ACCEPTANCE = 25&
Global Const M_CHAR_INVALID = 26&
Global Const M_TARGET_CHAR_SIZE_X = 27&
Global Const M_TARGET_CHAR_SIZE_Y = 28&
Global Const M_TARGET_CHAR_SPACING = 29&
Global Const M_DEBUG = 30&
Global Const M_FONT_TYPE = 31&
Global Const M_CHAR_NUMBER = 32&
Global Const M_CHAR_BOX_SIZE_X = 33&
Global Const M_CHAR_BOX_SIZE_Y = 34&
Global Const M_CHAR_OFFSET_X = 35&
Global Const M_CHAR_OFFSET_Y = 36&
Global Const M_CHAR_SIZE_X = 37&
Global Const M_CHAR_SIZE_Y = 38&
Global Const M_CHAR_THICKNESS = 39&
Global Const M_STRING_LENGTH = 40&
Global Const M_FONT_INIT_FLAG = 41&
Global Const M_CHAR_IN_FONT = 42&
Global Const M_ADD_TOP_HAT = 43&
Global Const M_KILL_BORDER = 44&
Global Const M_CHAR_ERASE = 45&
Global Const M_MODEL_LIST = 46&
Global Const M_CHAR_NUMBER_IN_FONT = 47&
Global Const M_STRING_VALIDATION = 48&
Global Const M_STRING_VALIDATION_HANDLER_PTR = M_STRING_VALIDATION
Global Const M_STRING_VALIDATION_HANDLER_USER_PTR = 49&
Global Const M_READ_SPEED = 50&
Global Const M_READ_ACCURACY = 51&
Global Const M_READ_FIRST_LEVEL = 52&
Global Const M_READ_LAST_LEVEL = 53&
Global Const M_READ_FAST_FIND = 55&
Global Const M_READ_ROBUSTNESS = 56&
Global Const M_STRING_LOC_NB_MODELS = 57&
Global Const M_READ_MODEL_STEP = 58&
Global Const M_PAT_ON_ACCELERATED = 59&
Global Const M_BLOB_ON_ACCELERATED = 60&
Global Const M_PROC_ON_ACCELERATED = 61&

Global Const M_CONSTRAINT = &H4000000
Global Const M_CONSTRAINT_TYPE = &H8000000

Global Const M_DISPLAY_ENABLE = 2&
Global Const M_BENCHMARK_ENABLE = 4&

'// MocrControl (M_SKIP_STRING_LOCATION possible values)
Global Const M_STRING_LOCATION_BLOB_ONLY = 2&
Global Const M_STRING_LOCATION_SEARCH_ONLY = 3&
Global Const M_STRING_LOCATION_BLOB_THAN_SEARCH = 4&
Global Const M_STRING_LOCATION_SEARCH_THAN_BLOB = 5&

'// MocrAllocFont()
Global Const M_SEMI_M12_92 = &H1&
Global Const M_SEMI_M13_88 = &H2&
Global Const M_SEMI = &H3&
Global Const M_FOREGROUND_WHITE = &H80&
Global Const M_FOREGROUND_BLACK = &H100&

'// MocrCalibrateFont()

'// MocrImportFont(), MocrRestoreFont(), MocrSaveFont()
Global Const M_LOAD_CONSTRAINT = &H2&
Global Const M_LOAD_CONTROL = &H4&
Global Const M_LOAD_CHARACTER = &H8&
Global Const M_SAVE = &H100&
Global Const M_SAVE_CONSTRAINT = &H200&
Global Const M_SAVE_CONTROL = &H400&
Global Const M_SAVE_CHARACTER = &H800&
Global Const M_FONT_MIL = &H8000&
Global Const M_FONT_ASCII = &H4000&

'// MocrGetResult()
Global Const M_STRING_VALID_FLAG = 1&
Global Const M_STRING_SCORE = M_SCORE
Global Const M_STRING = 3&
Global Const M_CHAR_VALID_FLAG = 4&
Global Const M_CHAR_SCORE = 5&
Global Const M_CHAR_POSITION_X = 6&
Global Const M_CHAR_POSITION_Y = 7&
Global Const M_GOOD_LOCATION_QUALITY_FLAG = 8&
Global Const M_CHAR_SIZE_SCORE = 9&
Global Const M_CHAR_MIN_OFFSET_X = 10&
Global Const M_CHAR_MAX_OFFSET_X = 11&
Global Const M_CHAR_MIN_OFFSET_Y = 12&
Global Const M_CHAR_MAX_OFFSET_Y = 13&

'// MocrCopyFont()
Global Const M_COPY_TO_FONT = 1&
Global Const M_COPY_FROM_FONT = 2&
Global Const M_CHARACTER_PAT_MODEL = &H10000
Global Const M_ALL_CHAR = &H8000&
Global Const M_SKIP_SEMI_STRING_UPDATE = &H4000&

'// MocrSetConstraint()
Global Const M_LETTER = &H2&
Global Const M_DIGIT = &H3&
Global Const M_UPPERCASE = &H10000
Global Const M_LOWERCASE = &H8000&

'// MocrModifyFont()
Global Const M_INVERT = 2&

'// MocrValidateString()
Global Const M_PRESENT = 2&
Global Const M_CHECK_VALID = 3&
Global Const M_CHECK_VALID_FAST = 4&

'******************************************************************************






'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Filename:  MILPROTO.BAS
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   1.0  $
'* Content :  This file contains the prototypes for the Matrox
'*            Imaging Library (MIL) C user's functions.
'* COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'******************************************************************************


'/***************************************************************************/
'/* BASIC IMAGE PROCESSING MODULE:                                          */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* POINT TO POINT : */

'/* -------------------------------------------------------------- */
Declare Sub MimArith Lib "milim.dll" (ByVal Src1ImageIdOrConstant As Double, ByVal Src2ImageIdOrConstant As Double, ByVal DestImageId As Long, ByVal Operation As Long)
Declare Sub MimArithMultiple Lib "milim.dll" (ByVal Src1ImageIdOrConstant As Double, ByVal Src2ImageIdOrConstant As Double, ByVal Src3ImageIdOrConstant As Double, ByVal Src4ImageIdOrConstant As Double, ByVal Src5ImageIdOrConstant As Double, ByVal DestImageId As Long, ByVal Operation As Long, ByVal OperationFlag As Long)
Declare Sub MimLutMap Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal LutId As Long)
Declare Sub MimShift Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbBitsToShift As Long)
Declare Sub MimBinarize Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal Condition As Long, ByVal CondLow As Double, ByVal CondHigh As Double)
Declare Sub MimClip Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal Condition As Long, ByVal CondLow As Double, ByVal CondHigh As Double, ByVal WriteLow As Double, ByVal WriteHigh As Double)
Declare Sub MimConvert Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal ConversionType As Long)

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* NEIGHBOURHOOD : */

'/* -------------------------------------------------------------- */
Declare Sub MimConvolve Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal KernelId As Long)
Declare Sub MimRank Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal KernelId As Long, ByVal Rank As Long, ByVal Mode As Long)
Declare Sub MimEdgeDetect Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestIntensityImageId As Long, ByVal DestAngleImageId As Long, ByVal KernelId As Long, ByVal ControlFlag As Long, ByVal ThresholdVal As Long)


'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* MORPHOLOGICAL: */

'/* -------------------------------------------------------------- */

Declare Sub MimLabel Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal Mode As Long)
Declare Sub MimConnectMap Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal LutBufId As Long)
Declare Sub MimDilate Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimErode Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimClose Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimOpen Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimMorphic Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal StructElementId As Long, ByVal Operation As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimThin Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimThick Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal NbIteration As Long, ByVal Mode As Long)
Declare Sub MimDistance Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal DistanceTranform As Long)
Declare Sub MimZoneOfInfluence Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal OperationFlag As Long)

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* GEOMETRICAL: */

'/* -------------------------------------------------------------- */
Declare Sub MimResize Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal FactorX As Double, ByVal FactorY As Double, ByVal InterpolationType As Long)
Declare Sub MimRotate Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DstImageId As Long, ByVal Angle As Double, ByVal SrcCenX As Double, ByVal SrcCenY As Double, ByVal DstCenX As Double, ByVal DstCenY As Double, ByVal InterpolationMode As Long)
Declare Sub MimTranslate Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestImageId As Long, ByVal xShift As Double, ByVal yShift As Double, ByVal InterpolationType As Long)
Declare Sub MimFlip Lib "milim.dll" (ByVal DestImageId As Long, ByVal SrcImageId As Long, ByVal Operation As Long, ByVal Mode As Long)

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* STATISTICAL: */

'/* -------------------------------------------------------------- */
Declare Sub MimHistogram Lib "milim.dll" (ByVal SrcImageId As Long, ByVal HistogramListId As Long)
Declare Sub MimHistogramEqualize Lib "milim.dll" (ByVal src_id As Long, ByVal dst_id As Long, ByVal EqualizationType As Long, ByVal Alpha As Double, ByVal Min As Double, ByVal Max As Double)
Declare Sub MimProject Lib "milim.dll" (ByVal SrcImageId As Long, ByVal DestArrayId As Long, ByVal ProjectionAngle As Double)
Declare Sub MimFindExtreme Lib "milim.dll" (ByVal SrcImageId As Long, ByVal ResultListId As Long, ByVal ExtremeType As Long)
Declare Sub MimLocateEvent Lib "milim.dll" (ByVal SrcImageId As Long, ByVal EventResultId As Long, ByVal Condition As Long, ByVal CondLow As Double, ByVal CondHigh As Double)
Declare Sub MimCountDifference Lib "milim.dll" (ByVal Src1ImageId As Long, ByVal Src2ImageId As Long, ByVal ResId As Long)
Declare Sub MimFree Lib "milim.dll" (ByVal ImResultId As Long)
Declare Sub MimGetResult1d Lib "milim.dll" (ByVal ImResultId As Long, ByVal Offresult As Long, ByVal Sizeresult As Long, ByVal ResultType As Long, UserTargetArrayPtr As Any)
Declare Sub MimGetResult Lib "milim.dll" (ByVal ImResultId As Long, ByVal ResultType As Long, UserTargetArrayPtr As Any)
Declare Function MimInquire Lib "milim.dll" (ByVal BufId As Long, ByVal InquireType As Long, TargetVarPtr As Any) As Long
Declare Function MimAllocResult Lib "milim.dll" (ByVal SystemId As Long, ByVal NumberOfResultElement As Long, ByVal ResultType As Long, IdVarPtr As Long) As Long


'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* TRANSFORM: */

'/* -------------------------------------------------------------- */


'/* -------------------------------------------------------------- */


'/***************************************************************************/
'/* GRAPHIC MODULE:                                                         */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */
Declare Function MgraAlloc Lib "mil.dll" (ByVal SystemId As Long, GraphContextIdVarPtr As Long) As Long
Declare Sub MgraFree Lib "mil.dll" (ByVal GraphContextId As Long)
Declare Sub MgraColor Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ForegroundColor As Double)
Declare Sub MgraBackColor Lib "mil.dll" (ByVal GraphContextId As Long, ByVal BackgroundColor As Double)
Declare Sub MgraFont Lib "milvb.dll" Alias "MgraFontInter" (ByVal GraphContextId As Long, ByVal Font1 As Long)
Declare Sub MgraFontScale Lib "mil.dll" (ByVal GraphContextId As Long, ByVal xFontScale As Double, ByVal yFontScale As Double)
Declare Sub MgraInquire Lib "mil.dll" (ByVal GraphContextId As Long, ByVal InquireType As Long, result_ptr As Any)
Declare Sub MgraControl Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ControlType As Long, ByVal ControlValue As Long)




'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* DRAWING : */

'/* -------------------------------------------------------------- */
Declare Sub MgraDot Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XPos As Long, ByVal YPos As Long)
Declare Sub MgraLine Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal XEnd As Long, ByVal YEnd As Long)
Declare Sub MgraArc Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XCenter As Long, ByVal YCenter As Long, ByVal XRad As Long, ByVal YRad As Long, ByVal StartAngle As Double, ByVal EndAngle As Double)
Declare Sub MgraArcFill Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XCenter As Long, ByVal YCenter As Long, ByVal XRad As Long, ByVal YRad As Long, ByVal StartAngle As Double, ByVal EndAngle As Double)
Declare Sub MgraRect Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal XEnd As Long, ByVal YEnd As Long)
Declare Sub MgraRectFill Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal XEnd As Long, ByVal YEnd As Long)
Declare Sub MgraFill Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XStart As Long, ByVal YStart As Long)
Declare Sub MgraClear Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long)
Declare Sub MgraText Lib "mil.dll" (ByVal GraphContextId As Long, ByVal ImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal String1 As String)

'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* DATA GENERATION MODULE:                                                 */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* DATA BUFFERS: */

'/* -------------------------------------------------------------- */
Declare Sub MgenLutRamp Lib "mil.dll" (ByVal LutBufId As Long, ByVal StartPoint As Long, ByVal StartValue As Double, ByVal EndPoint As Long, ByVal EndValue As Double)
Declare Sub MgenLutFunction Lib "mil.dll" (ByVal lut_id As Long, ByVal func As Long, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal start_index As Long, ByVal StartXValue As Double, ByVal end_index As Long)

'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* DATA BUFFERS MODULE:                                                    */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CREATION: */

'/* -------------------------------------------------------------- */
Declare Function MbufAlloc1d Lib "mil.dll" (ByVal SystemId As Long, ByVal SizeX As Long, ByVal Type1 As Long, ByVal BufAttribute As Long, IdVarPtr As Long) As Long
Declare Function MbufAlloc2d Lib "mil.dll" (ByVal SystemId As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal Type1 As Long, ByVal BufAttribute As Long, IdVarPtr As Long) As Long
Declare Function MbufAllocColor Lib "mil.dll" (ByVal SystemId As Long, ByVal SizeBand As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal Type1 As Long, ByVal BufAttribute As Long, IdVarPtr As Long) As Long
Declare Function MbufChild1d Lib "mil.dll" (ByVal ParentImageId As Long, ByVal OffX As Long, ByVal SizeX As Long, IdVarPtr As Long) As Long
Declare Function MbufChild2d Lib "mil.dll" (ByVal ParentMilBufId As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, IdVarPtr As Long) As Long
Declare Function MbufChildColor Lib "mil.dll" (ByVal ParentMilBufId As Long, ByVal Band As Long, IdVarPtr As Long) As Long
Declare Function MbufCreateColor Lib "mil.dll" (ByVal SystemId As Long, ByVal SizeBand As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal BufferType As Long, ByVal BufAttribute As Long, ByVal ControlFlag As Long, ByVal Pitch As Long, ArrayOfDataPtr As Any, IdVarPtr As Long) As Long
Declare Sub MbufFree Lib "mil.dll" (ByVal BufId As Long)
Declare Function MbufChildColor2d Lib "mil.dll" (ByVal ParentMilBufId As Long, ByVal Band As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, IdVarPtr As Long) As Long
Declare Function MbufCreate2d Lib "mil.dll" (ByVal ParentMilBufId As Long, SizeX As Long, ByVal SizeY As Long, ByVal typ As Long, ByVal Attrib As Long, ByVal ControlFlag As Long, ByVal Pitch As Long, ArrayOfDataPtr As Any, IdVarPtr As Long) As Long
                                              
                               
'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* ACCESS: */

'/* -------------------------------------------------------------- */
Declare Sub MbufClear Lib "mil.dll" (ByVal BufId As Long, ByVal Value As Double)
Declare Sub MbufCopy Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long)
Declare Sub MbufCopyColor Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long, ByVal Band As Long)
Declare Sub MbufCopyClip Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long, ByVal DestOffsetX As Long, ByVal DestOffsetY As Long)
Declare Sub MbufCopyMask Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long, ByVal MaskValue As Long)
Declare Sub MbufCopyCond Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long, ByVal CondBufId As Long, ByVal Cond As Long, ByVal CondVal As Double)
Declare Sub MbufPut1d Lib "mil.dll" (ByVal DestBufId As Long, ByVal OffX As Long, ByVal SizeX As Long, BufferPtr As Any)
Declare Sub MbufPut2d Lib "mil.dll" (ByVal DestBufId As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, BufferPtr As Any)
Declare Sub MbufPutColor Lib "mil.dll" (ByVal DestBufId As Long, ByVal Format1 As Long, ByVal Band As Long, BufferPtr As Any)
Declare Sub MbufPut Lib "mil.dll" (ByVal DestBufId As Long, BufferPtr As Any)
Declare Sub MbufGet1d Lib "mil.dll" (ByVal SourceBufId As Long, ByVal OffX As Long, ByVal SizeX As Long, BufferPtr As Any)
Declare Sub MbufGet2d Lib "mil.dll" (ByVal SourceBufId As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, BufferPtr As Any)
Declare Sub MbufGetColor Lib "mil.dll" (ByVal SourceBufId As Long, ByVal Format1 As Long, ByVal Band As Long, BufferPtr As Any)
Declare Sub MbufGet Lib "mil.dll" (ByVal SourceBufId As Long, BufferPtr As Any)
Declare Sub MbufSave Lib "mil.dll" (ByVal FileName As String, ByVal BufId As Long)
Declare Sub MbufLoad Lib "mil.dll" (ByVal FileName As String, ByVal BufId As Long)
Declare Sub MbufExport Lib "mil.dll" (ByVal FileName As String, ByVal FileFormatId As Long, ByVal srcBufId As Long)
Declare Sub MbufGetLine Lib "mil.dll" (ByVal SrcImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal XEnd As Long, ByVal YEnd As Long, ByVal Mode As Long, NbPixelsValPtr As Long, BufferType As Any)
Declare Sub MbufPutLine Lib "mil.dll" (ByVal SrcImageId As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal XEnd As Long, ByVal YEnd As Long, ByVal Mode As Long, NbPixelsValPtr As Long, BufferType As Any)
Declare Sub MappChild Lib "mil.dll" (ByVal ParentId As Long, ByVal InitFlag As Long, IdVarPtr As Long)
Declare Sub MappTimer Lib "mil.dll" (ByVal Mode As Long, Time As Double)
Declare Sub MbufModified2d Lib "mil.dll" (ByVal BufferId As Long, ByVal OffsetX As Long, ByVal OffsetY As Long, ByVal SizeX As Long, ByVal SizeY As Long)
Declare Function MbufInquire Lib "mil.dll" (ByVal BufId As Long, ByVal InquireType As Long, ResultPtr As Any) As Long
Declare Function MbufDiskInquire Lib "mil.dll" (ByVal FileName As String, ByVal InquireType As Long, ResultPtr As Any) As Long
Declare Function MbufRestore Lib "mil.dll" (ByVal FileName As String, ByVal SystemId As Long, IdVarPtr As Long) As Long
Declare Function MbufImport Lib "mil.dll" (ByVal FileName As String, ByVal FileFormatId As Long, ByVal Operation As Long, ByVal SystemId As Long, IdVarPtr As Long) As Long
Declare Function MappControlThread Lib "mil.dll" (ByVal ThreadOrEventId As Long, ByVal Operation As Long, ByVal OperationValue As Long, IdVarPtr As Long) As Long
Declare Sub MbufCopyColor2d Lib "mil.dll" (ByVal srcBufId As Long, ByVal DestBufId As Long, ByVal SrcBand As Long, ByVal SrcOffX As Long, ByVal SrcOffY As Long, ByVal DstBand As Long, ByVal DstOffX As Long, ByVal DstOffY As Long, ByVal SizeX As Long, ByVal SizeY As Long)
Declare Sub MbufPutColor2d Lib "mil.dll" (ByVal DestBufId As Long, ByVal Format As Long, ByVal Band As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, BufferPtr As Any)
Declare Sub MbufGetColor2d Lib "mil.dll" (ByVal SourceBufId As Long, ByVal Format As Long, ByVal Band As Long, ByVal OffX As Long, ByVal OffY As Long, ByVal SizeX As Long, ByVal SizeY As Long, BufferPtr As Any)

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */
Declare Sub MbufControlNeighborhood Lib "mil.dll" (ByVal BufId As Long, ByVal OperationFlags As Long, ByVal OperationValue As Long)
Declare Sub MbufControl Lib "mil.dll" (ByVal BufId As Long, ByVal OperationFlags As Long, ByVal OperationValue As Double)


'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* I/O DEVICES:                                                            */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CREATION: */

'/* -------------------------------------------------------------- */
Declare Function MdigAlloc Lib "mil.dll" (ByVal SystemId As Long, ByVal DeviceNum As Long, ByVal DataFormat As String, ByVal InitFlag As Long, IdVarPtr As Long) As Long
Declare Sub MdigFree Lib "mil.dll" (ByVal DevId As Long)


'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */

Declare Sub MdigChannel Lib "mil.dll" (ByVal DevId As Long, ByVal Channel As Long)
Declare Sub MdigReference Lib "mil.dll" (ByVal DevId As Long, ByVal ReferenceType As Long, ByVal ReferenceLevel As Long)
Declare Sub MdigLut Lib "mil.dll" (ByVal DevId As Long, ByVal LutBufId As Long)
Declare Sub MdigHalt Lib "mil.dll" (ByVal DevId As Long)
Declare Function MdigInquire Lib "mil.dll" (ByVal DevId As Long, ByVal InquireType As Long, ResultPtr As Any) As Long
Declare Function MdigInquireString Lib "mil.dll" Alias "MdigInquire" (ByVal DevId As Long, ByVal InquireType As Long, ByVal ResultPtr As String) As Long
Declare Sub MdigControl Lib "mil.dll" (ByVal DigitizerId As Long, ByVal ControlType As Long, ByVal Value As Double)
Declare Sub MdigGrabWait Lib "mil.dll" (ByVal DevId As Long, ByVal Flag As Long)
Declare Function MdigHookFunction Lib "mil.dll" (ByVal SrcDevId As Long, ByVal HookType As Long, ByVal HookHandlerPtr As Any, ByRef UserDataPtr As Any) As Long


'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* ACCESS: */

'/* -------------------------------------------------------------- */

Declare Sub MdigGrab Lib "mil.dll" (ByVal SrcDevId As Long, ByVal DestImageId As Long)
Declare Sub MdigGrabContinuous Lib "mil.dll" (ByVal SrcDevId As Long, ByVal DestImageId As Long)
Declare Sub MdigAverage Lib "mil.dll" (ByVal Digitizer As Long, ByVal DestImageId As Long, ByVal WeightFactor As Long, ByVal AverageType As Long, ByVal NbIteration As Long)

'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* DISPLAY MODULE:                                                         */
'/***************************************************************************/

'/* -------------------------------------------------------------- */
'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */
Declare Function MdispAlloc Lib "mil.dll" (ByVal SystemId As Long, ByVal DispNum As Long, ByVal DispFormat As String, ByVal InitFlag As Long, IdVarPtr As Long) As Long
Declare Function MdispInquire Lib "mil.dll" (ByVal DisplayId As Long, ByVal inquire_type As Long, result_ptr As Any) As Long
Declare Sub MdispFree Lib "mil.dll" (ByVal DisplayId As Long)
Declare Sub MdispSelect Lib "mil.dll" (ByVal DisplayId As Long, ByVal ImageId As Long)
Declare Sub MdispDeselect Lib "mil.dll" (ByVal DisplayId As Long, ByVal ImageId As Long)
Declare Sub MdispPan Lib "mil.dll" (ByVal DisplayId As Long, ByVal XOffset As Long, ByVal YOffset As Long)
Declare Sub MdispZoom Lib "mil.dll" (ByVal DisplayId As Long, ByVal XFactor As Long, ByVal YFactor As Long)
Declare Sub MdispLut Lib "mil.dll" (ByVal DisplayId As Long, ByVal LutBufId As Long)
Declare Sub MdispOverlayKey Lib "mil.dll" (ByVal DisplayId As Long, ByVal Mode As Long, ByVal Cond As Long, ByVal Mask As Long, ByVal Color As Long)
Declare Sub MdispControl Lib "mil.dll" (ByVal DisplayId As Long, ByVal ControlType As Long, ByVal Value As Long)
Declare Function MdispHookFunction Lib "mil.dll" (ByVal SrcDevId As Long, ByVal HookType As Long, ByVal HookHandlerPtr As Any, ByRef UserDataPtr As Any) As Long

'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* SYSTEM MODULE:                                                          */
'/***************************************************************************/

'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */
Declare Function MsysInquire Lib "mil.dll" (ByVal SystemId As Long, ByVal InquireType As Long, ResultPtr As Any) As Long
Declare Function MsysAlloc Lib "milvb.dll" Alias "MsysAllocInter" (ByVal SystemType As Long, ByVal SystemNum As Long, ByVal InitFlag As Long, IdVarPtr As Long) As Long
Declare Sub MsysFree Lib "milvb.dll" Alias "MsysFreeInter" (ByVal SystemId As Long)
Declare Sub MsysControl Lib "mil.dll" (ByVal SystemId As Long, ByVal ControlType As Long, ByVal TargetSysId As Long)
Declare Sub MsysConfigAccess Lib "mil.dll" (ByVal SystemId As Long, ByVal VendorId As Long, ByVal DeviceId As Long, ByVal DeviceNum As Long, ByVal OperationFlag As Long, ByVal OperationType As Long, ByVal Offset As Long, ByVal Size As Long, UserArrayPtr As Any)

'/* -------------------------------------------------------------- */

'/***************************************************************************/
'/* APPLICATION MODULE:                                                     */
'/***************************************************************************/

'/* -------------------------------------------------------------- */

'/* CONTROL: */

'/* -------------------------------------------------------------- */
Declare Function MappAlloc Lib "mil.dll" (ByVal InitFlag As Long, IdVarPtr As Long) As Long
Declare Function MappGetError Lib "mil.dll" (ByVal ErrorType As Long, ErrorVarPtr As Any) As Long
Declare Function MappGetErrorString Lib "mil.dll" Alias "MappGetError" (ByVal ErrorType As Long, ByVal ErrorVarPtr As String) As Long
Declare Function MappGetHookInfo Lib "mil.dll" (ByVal Id As Long, ByVal InfoType As Long, UserPtr As Any) As Long
Declare Function MappInquire Lib "mil.dll" (ByVal InquireType As Long, UserVarPtr As Any) As Long
Declare Sub MappFree Lib "mil.dll" (ByVal ApplicationId As Long)
Declare Sub MappControl Lib "mil.dll" (ByVal ControlType As Long, ByVal ControlFlag As Long)
Declare Sub MappModify Lib "mil.dll" (ByVal FirstId As Long, ByVal SecondId As Long, ByVal ModificationType As Long, ByVal ModificationFlag As Long)
Declare Function MappHookFunction Lib "mil.dll" (ByVal HookType As Long, ByVal HookHandlerPtr As Any, ByRef UserDataPtr As Any) As Long


                                           
'/* -------------------------------------------------------------- */




'/**************************************************************************/
'/* VGA WINDOWS SPECIFIC FUNCTION SET                                              */
'/**************************************************************************/
Declare Sub MvgaDispSelectClientArea Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal MilVgaBufferId As Long, ByVal ClientWindowHandle As Long)
Declare Sub MvgaDispDeselectClientArea Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal MilVgaBufferId As Long, ByVal ClientWindowHandle As Long)
Declare Sub MvgaDispControl Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal ControlType As Long, ByVal ControlState As Long)
Declare Sub MvgaDispCapture Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal OffsetX As Long, ByVal OffsetY As Long, ByVal SizeX As Long, ByVal SizeY As Long, ByVal DestOffsetX As Long, ByVal DestOffsetY As Long)
Declare Sub MvgaDispProtectArea Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal OffsetX As Long, ByVal OffsetY As Long, ByVal SizeX As Long, ByVal SizeY As Long)
Declare Sub MvgaDispSetTitleName Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal TitleName As String)
Declare Sub MdispSelectWindow Lib "mil.dll" (ByVal DisplayId As Long, ByVal ImageId As Long, ByVal ClientWindowHandle As Long)
Declare Function MvgaDispInquire Lib "mil.dll" (ByVal MilVgaDisplayId As Long, ByVal InquireType As Long, TargetVarPtr As Long) As Long


'******************************************************************************






'******************************************************************************
'******************************************************************************
'******************************************************************************
'*
'*    Filename:  GENESIS.BAS
'*    Owner   :  Matrox Imaging dept.
'*    Rev     :  $Revision:   4.0  $
'*    Content :  This file contains the new defines that are needed by the user
'*               to use the MIL current library with the GENESIS.
'*
'*    COPYRIGHT (c) Matrox Electronic Systems Ltd.
'*    All Rights Reserved
'*
'******************************************************************************
'******************************************************************************
'******************************************************************************

'*****************************************************
'*          GENESIS DRIVER RELEASE (4.00)
'*
'*     These are the new defines for the GENESIS
'*****************************************************

'*******************************************************************
'*
'* COPYRIGHT (c) 1994-1996 Matrox Electronic Systems Ltd.
'* All Rights Reserved
'*
'*******************************************************************

'* Nothing for now.

'*************************************************************************





'******************************************************************************
'******************************************************************************
'******************************************************************************
'* Filename:  METEOR.H
'* Owner   :  Matrox Imaging dept.
'* Rev     :  $Revision:   4.0  $
'* Content :  This file contains the defines that are needed by the user
'*            to use the MIL library with the Meteor.
'* COPYRIGHT (c) Matrox Electronic Systems Ltd.
'* All Rights Reserved
'******************************************************************************
'******************************************************************************
'******************************************************************************

'/******************************************************************************/
'/*                    METEOR DRIVER RELEASE (4.00)                            */
'/*                                                                            */
'/*    These are the specifics or not yet released defines for the METEOR      */
'/******************************************************************************/
Global Const M_IRQ_GLOBAL_OBJECT = 125&

Global Const M_RS170 = &H1&
Global Const M_NTSC = &H2&
Global Const M_CCIR = &H3&
Global Const M_PAL = &H4&
Global Const M_NTSC_RGB = &H5&
Global Const M_PAL_RGB = &H6&
Global Const M_SECAM_RGB = &H7&
Global Const M_SECAM = &H8&

Global Const M_RGB888_ATIMACH64 = &HE5E5E500       '/* Internal use only */
Global Const M_RGB888_NORMALVGA = &H39393900       '/* Internal use only */
Global Const M_SET_ROUTER_TO_CH0 = &HFFFFFFC3           '/* Internal use only */
Global Const M_SET_ROUTER_TO_CH1 = &HAAAAAAC3           '/* Internal use only */
Global Const M_SET_ROUTER_TO_CH2 = &H555555C3      '/* Internal use only */

Global Const M_TUNER_STANDARD_L = 0&
Global Const M_TUNER_STANDARD_L_PRIME = 1&
Global Const M_TUNER_STANDARD_BG = 2&
Global Const M_TUNER_STANDARD_MN = 3&

'/* MTLLioWrite() and MTLLioRead() defines */
'/******************************************************************************/
Global Const M_FI1200 = &HC000&
Global Const M_PCF8574_0 = &H7000&
Global Const M_PCF8574_1 = &H7200&
Global Const M_TDA9855 = &HB600&
Global Const M_SAA7196 = &H4000&
Global Const M_SAA7116 = &H0&
Global Const M_BT254 = &H1&

'/* User Hook identification*/
'/******************************************************************************/
Global Const ISR_START_OF_FIELD_MASK = &HFF&
Global Const ISR_START_OF_FIELD_BIT = &H1&
Global Const ISR_START_OF_FIELD_ODD_BIT = &H2&
Global Const ISR_START_OF_FIELD_EVEN_BIT = &H4&
Global Const ISR_START_OF_FRAME_BIT = &H8&
Global Const ISR_START_OF_GRAB_BIT = &H10&
Global Const ISR_START_OF_GRAB_FRAME_BIT = &H20&

Global Const ISR_END_OF_FIELD_MASK = &HFF00&
Global Const ISR_END_OF_FIELD_EVEN_BIT = &H100&
Global Const ISR_END_OF_FIELD_ODD_BIT = &H200&
Global Const ISR_END_OF_FIELD_BIT = &H400&
Global Const ISR_END_OF_GRAB_FRAME_BIT = &H800&

Global Const ISR_END_OF_GRAB_MASK = (&HFF0000 Or ISR_END_OF_GRAB_FRAME_BIT)              '// special case for linescan no sync
Global Const ISR_END_OF_GRAB_BIT = &H10000
                                        

'/* PCI device information                                                     */
'/******************************************************************************/
Global Const M_PCI_VENDOR_ID = &H0               '// (16 lsb)
Global Const M_PCI_DEVICE_ID = &H0               '// (16 msb)
Global Const M_PCI_COMMAND = &H1                 '// (16 lsb)
Global Const M_PCI_STATUS = &H1                  '// (16 msb)
Global Const M_PCI_REVISION_ID = &H2             '// (byte 0)
Global Const M_PCI_CLASS_CODE = &H2              '// (byte 1,2,3)
Global Const M_PCI_LATENCY_TIMER = &H3           '// (byte 1)
Global Const M_PCI_BASE_ADRS0 = &H4              '//
Global Const M_PCI_BASE_ADRS1 = &H5              '//
Global Const M_PCI_INT_LINE = &HF                '// (byte 0)
Global Const M_PCI_INT_PIN = &HF                 '// (byte 1)


'******************************************************************************



'******************************************************************************
'******************************************************************************
'******************************************************************************
'*
'*
'*    Filename:  PULSAR.BAS
'*    Owner   :  Matrox Imaging dept.
'*    Rev     :  $Revision:   4.0  $
'*    Content :  This file contains the defines that are needed by the user
'*               to use the MIL library with the PULSAR.
'*
'*    COPYRIGHT (c) Matrox Electronic Systems Ltd.
'*    All Rights Reserved
'*
'******************************************************************************
'******************************************************************************
'******************************************************************************

'* Nothing for now.

'******************************************************************************





'******************************************************************************
'******************************************************************************
'******************************************************************************
'/* Filename:  MILSETUP.H                                                  */
'/* Owner   :  Matrox Imaging dept.                                        */
'/* Rev     :  $Revision:   1.0  $                                         */
'/* Content :  This file contains definitions for specifying the target    */
'/*            compile environment and the default state to set for        */
'/*            MIL (Matrox Imaging Library). It also defines the           */
'/*            MappAllocDefault() and MappFreeDefault() macros.            */
'/* COPYRIGHT (c) 1993-1995  Matrox Electronic Systems Ltd.                */
'/* All Rights Reserved.                                                   */
'******************************************************************************
'******************************************************************************
'******************************************************************************



'* MIL directory
Global Const M_MIL_DIRECTORY = "c:\\mil"


'/************************************************************************/
'/* MIL LITE IDENTIFICATION FLAG                                         */
'/* Activate or Deactivate MIL Lite flag                                 */
'/************************************************************************/

Global Const M_MIL_LITE = 0

'/************************************************************************/
'/* SETUP SPECIFIED FLAG                                                 */
'/* Activate or Deactivate MIL use-setup flag                            */
'/************************************************************************/

Global Const M_MIL_USE_SETUP = 1

'/************************************************************************/
'/* COMPILATION FLAG                                                     */
'/* One and only one flag must be active                                 */
'/*                                                                      */
'/* Activate or Deactivate DOS compile mode                              */
'/* Activate or Deactivate WINDOWS compile mode                          */
'/* Activate or Deactivate PHARLAP LIB for LIB compile mode              */
'/* Activate or Deactivate DOS32 compile mode                            */
'/* Activate or Deactivate NT under DOS compile mode                     */
'/* Activate or Deactivate NT under WINDOWS compile mode                 */
'/************************************************************************/

Global Const M_MIL_USE_OS = 1&
Global Const M_MIL_USE_DOS = 0&
Global Const M_MIL_USE_WINDOWS = 0&
Global Const M_MIL_USE_PHARLAP_LIB = 0&
Global Const M_MIL_USE_DOS_32 = 0&
Global Const M_MIL_USE_NT_DOS = 0&
Global Const M_MIL_USE_NT_WINDOWS = 1&

'/************************************************************************/
'/* ERROR MESSAGES AND VALUES USAGE FLAG                                 */
'/* Activate or Deactivate error messages inclusion                      */
'/************************************************************************/

Global Const M_MIL_USE_ERROR_MESSAGE = 0&

'/************************************************************************/
'/* MMX USAGE FLAG (Internal use only)                                   */
'/************************************************************************/

Global Const M_MIL_USE_MMX = 1&

'/************************************************************************/
'/* BLOB ANALYSIS MODULE PROGRAMMING FLAG                                */
'/* Activate or Deactivate blob module programming                       */
'/************************************************************************/

Global Const M_MIL_USE_BLOB = 1&

'/************************************************************************/
'/* PATTERN RECOGNITION MODULE PROGRAMMING FLAG                          */
'/* Activate or Deactivate pat module programming                        */
'/************************************************************************/

Global Const M_MIL_USE_PAT = 1&

'/************************************************************************/
'/* OCR MODULE PROGRAMMING FLAG                                          */
'/* Activate or Deactivate ocr module programming                        */
'/************************************************************************/

Global Const M_MIL_USE_OCR = 1&

'/************************************************************************/
'/* MEASUREMENT MODULE PROGRAMMING FLAG                                  */
'/* Activate or Deactivate meas module programming                       */
'/************************************************************************/

Global Const M_MIL_USE_MEAS = 1&

'/************************************************************************/
'/* NATIVE MODE PROGRAMMING FLAG                                         */
'/* Activate or Deactivate native mode programming                       */
'/************************************************************************/

Global Const M_MIL_USE_NATIVE = 1&

'/************************************************************************/
'/* METEOR SYSTEM USAGE FLAG AND INCLUDE PATH                            */
'/* Activate or Deactivate meteor module programming                     */
'/************************************************************************/

Global Const M_MIL_USE_METEOR = 1&

'/************************************************************************/
'/* PULSAR SYSTEM USAGE FLAG AND INCLUDE PATH                            */
'/* Activate or Deactivate pulsar module programming                     */
'/************************************************************************/

Global Const M_MIL_USE_PULSAR = 1&

'/************************************************************************/
'/* GENESIS SYSTEM USAGE FLAG AND INCLUDE PATH                           */
'/* Activate or Deactivate Genesis module programming                    */
'/************************************************************************/

Global Const M_MIL_USE_GENESIS = 1&

'/************************************************************************/
'/* CORONA SYSTEM USAGE FLAG AND INCLUDE PATH                            */
'/* Activate or Deactivate Corona module programming                     */
'/************************************************************************/

Global Const M_MIL_USE_CORONA = 1&

'/************************************************************************/
'/* NEWBOARD SYSTEM USAGE FLAG AND INCLUDE PATH                          */
'/* Activate or Deactivate newboard module programming                   */
'/************************************************************************/

Global Const M_MIL_USE_NEWBOARD = 1&

'/************************************************************************/
'/* DEFAULT STATE INITIALIZATION FLAG                                    */
'/************************************************************************/

Global Const M_SETUP = M_COMPLETE

'/************************************************************************/
'/* DEFAULT SYSTEM SPECIFICATIONS                                        */
'/************************************************************************/

Global Const M_DEF_SYSTEM_TYPE = M_SYSTEM_VGA
Global Const M_DEF_SYSTEM_NUM = M_DEV0
Global Const M_SYSTEM_SETUP = M_DEF_SYSTEM_TYPE

'/************************************************************************/
'/* DEFAULT DIGITIZER SPECIFICATIONS                                     */
'/************************************************************************/

Global Const M_DEF_DIGITIZER_NUM = M_DEV0
Global Const M_DEF_DIGITIZER_FORMAT = "M_DEFAULT"
Global Const M_DEF_DIGITIZER_INIT = M_DEFAULT
Global Const M_CAMERA_SETUP = M_DEF_DIGITIZER_FORMAT

'/************************************************************************/
'/* DEFAULT DISPLAY SPECIFICATIONS                                       */
'/************************************************************************/

Global Const M_DEF_DISPLAY_NUM = M_DEV0
Global Const M_DEF_DISPLAY_FORMAT = "M_DEFAULT"
Global Const M_DEF_DISPLAY_INIT = M_DEFAULT
Global Const M_DISPLAY_SETUP = M_DEF_DISPLAY_FORMAT
Global Const M_DEF_DISPLAY_KEY_COLOR = 0
Global Const M_DEF_DISPLAY_KEY_ENABLE_ON_ALLOC = 0
Global Const M_DEF_DISPLAY_KEY_DISABLE_ON_FREE = 0

'/************************************************************************/
'/* DEFAULT IMAGE BUFFER SPECIFICATIONS                                  */
'/************************************************************************/

Global Const M_DEF_IMAGE_NUMBANDS_MIN = 1&
Global Const M_DEF_IMAGE_SIZE_X_MIN = 512
Global Const M_DEF_IMAGE_SIZE_Y_MIN = 480
Global Const M_DEF_IMAGE_SIZE_X_MAX = 1024
Global Const M_DEF_IMAGE_SIZE_Y_MAX = 1024
Global Const M_DEF_IMAGE_TYPE = 8 + M_UNSIGNED
Global Const M_DEF_IMAGE_ATTRIBUTE_MIN = M_IMAGE + M_PROC


'***************************************************************************
'* GLOBAL DECLARATION FOR MAPPALLOCDEFAULT IN VBASIC                       *
'***************************************************************************

Global Const M_NUL = -1
Global Const M_NO_ALLOC = -1



'/***************************************************************************/
'/* LocalBufferAllocDefault - Local macro to allocate a default MIL buffer: */
'/*                                                                         */
'/* MIL_ID *SystemIdVarPtr;                                                 */
'/* MIL_ID *DisplayIdVarPtr;                                                */
'/* MIL_ID *ImageIdVarPtr;                                                  */
'/*                                                                         */
'/***************************************************************************/
Sub LocalBufferAllocDefault(SystemIdVarPtr As Long, DisplayIdVarPtr As Long, DigitizerIdVarPtr As Long, ImageIdVarPtr As Long)
                                                                            
   '* local variables
   Dim m_def_image_numbands  As Long
   Dim m_def_image_size_x    As Long
   Dim m_def_image_size_y    As Long
   Dim m_def_image_attribute As Long
   Dim m_def_image_type_tmp  As Long

   m_def_image_numbands = M_DEF_IMAGE_NUMBANDS_MIN
   m_def_image_size_x = M_DEF_IMAGE_SIZE_X_MIN
   m_def_image_size_y = M_DEF_IMAGE_SIZE_Y_MIN
   m_def_image_type_tmp = M_DEF_IMAGE_TYPE
   m_def_image_attribute = M_DEF_IMAGE_ATTRIBUTE_MIN


   '/* determines the needed size band, x, y, type and attribute */
   If Not (DisplayIdVarPtr = M_NO_ALLOC) Then
      If (MdispInquire(DisplayIdVarPtr, M_DISPLAY_MODE, M_NULL) = M_WINDOWED) Then
         If Not (DigitizerIdVarPtr = M_NO_ALLOC) Then
            m_def_image_size_x = MdigInquire(DigitizerIdVarPtr, M_SIZE_X, M_NULL)
            m_def_image_size_y = MdigInquire(DigitizerIdVarPtr, M_SIZE_Y, M_NULL)
            m_def_image_type_tmp = MdigInquire(DigitizerIdVarPtr, M_TYPE, M_NULL)
         Else
            m_def_image_size_x = M_DEF_IMAGE_SIZE_X_MIN
            m_def_image_size_y = M_DEF_IMAGE_SIZE_Y_MIN
            m_def_image_type_tmp = M_DEF_IMAGE_TYPE
         End If
      Else
         m_def_image_size_x = MdispInquire(DisplayIdVarPtr, M_SIZE_X, M_NULL)
         m_def_image_size_y = MdispInquire(DisplayIdVarPtr, M_SIZE_Y, M_NULL)
         m_def_image_type_tmp = MdispInquire(DisplayIdVarPtr, M_TYPE, M_NULL)
      End If
   End If
   
   If (m_def_image_size_x < M_DEF_IMAGE_SIZE_X_MIN) Then
      m_def_image_size_x = M_DEF_IMAGE_SIZE_X_MIN
   End If
   If (m_def_image_size_y < M_DEF_IMAGE_SIZE_Y_MIN) Then
      m_def_image_size_y = M_DEF_IMAGE_SIZE_Y_MIN
   End If
   If (m_def_image_size_x > M_DEF_IMAGE_SIZE_X_MAX) Then
      m_def_image_size_x = M_DEF_IMAGE_SIZE_X_MAX
   End If
   If (m_def_image_size_y > M_DEF_IMAGE_SIZE_Y_MAX) Then
      m_def_image_size_y = M_DEF_IMAGE_SIZE_Y_MAX
   End If
   If (((m_def_image_type_tmp & &HFF) < (M_DEF_IMAGE_TYPE & &HFF)) Or ((m_def_image_type_tmp & &HFF) > (M_DEF_IMAGE_TYPE & &HFF))) Then
       m_def_image_type_tmp = M_DEF_IMAGE_TYPE
   End If
   
   '/* determines the needed attribute and number of band */
   m_def_image_attribute = M_DEF_IMAGE_ATTRIBUTE_MIN
   m_def_image_numbands = M_DEF_IMAGE_NUMBANDS_MIN
   If Not (DisplayIdVarPtr = M_NO_ALLOC) Then
      m_def_image_attribute = M_DISP + m_def_image_attribute
   End If
   If Not (DigitizerIdVarPtr = M_NO_ALLOC) Then
       m_def_image_attribute = M_GRAB + m_def_image_attribute
       m_def_image_numbands = MdigInquire(DigitizerIdVarPtr, M_SIZE_BAND, M_NULL)
       If (m_def_image_numbands < M_DEF_IMAGE_NUMBANDS_MIN) Then
         m_def_image_numbands = M_DEF_IMAGE_NUMBANDS_MIN
           End If
   End If

       
   '/* allocates a monochromatic or color image buffer */
   a = MbufAllocColor(SystemIdVarPtr, m_def_image_numbands, m_def_image_size_x, m_def_image_size_y, m_def_image_type_tmp, m_def_image_attribute, ImageIdVarPtr)
                                               

  '* clear and display the image buffer
  If (Not (DisplayIdVarPtr = M_NO_ALLOC) And Not (ImageIdVarPtr = M_NO_ALLOC)) Then
     Call MbufClear(ImageIdVarPtr, 0)
     Call MdispSelect(DisplayIdVarPtr, ImageIdVarPtr)
  End If
     
End Sub



'/**************************************************************************/
'/* MappAllocDefault - macro to allocate default MIL objects:              */
'/*                                                                        */
'/* long    InitFlag;                                                      */
'/* MIL_ID *ApplicationIdVarPtr;                                           */
'/* MIL_ID *SystemIdVarPtr;                                                */
'/* MIL_ID *DisplayIdVarPtr;                                               */
'/* MIL_ID *DigitizerIdVarPtr;                                             */
'/* MIL_ID *ImageIdVarPtr;                                                 */
'/*                                                                        */
'/* Note:                                                                  */
'/*       An application must be allocated before a system.                */
'/*       An system must be allocated before a display,digitzer or image.  */
'/*                                                                        */
'/**************************************************************************/


Sub MappAllocDefault(InitFlag As Long, ApplicationIdVarPtr As Long, SystemIdVarPtr As Long, DisplayIdVarPtr As Long, DigitizerIdVarPtr As Long, ImageIdVarPtr As Long)
                                                              
  '* allocate a MIL application.
  If Not (ApplicationIdVarPtr = M_NO_ALLOC) Then
     a = MappAlloc(InitFlag, ApplicationIdVarPtr)
  End If
                                                                       
  '* allocate a system
  If (Not (SystemIdVarPtr = M_NO_ALLOC) And Not (ApplicationIdVarPtr = M_NO_ALLOC)) Then
     a = MsysAlloc(M_DEF_SYSTEM_TYPE, M_DEF_SYSTEM_NUM, InitFlag, SystemIdVarPtr)
  End If
                                               
  '* allocate a display
  If (Not (DisplayIdVarPtr = M_NO_ALLOC) And Not (SystemIdVarPtr = M_NO_ALLOC)) Then
     a = MdispAlloc(SystemIdVarPtr, M_DEF_DISPLAY_NUM, M_DEF_DISPLAY_FORMAT, M_DEF_DISPLAY_INIT, DisplayIdVarPtr)
  End If
                                               
  '* allocate a digitizer
  If (Not (DigitizerIdVarPtr = M_NO_ALLOC) And Not (SystemIdVarPtr = M_NO_ALLOC)) Then
     a = MdigAlloc(SystemIdVarPtr, M_DEF_DIGITIZER_NUM, M_DEF_DIGITIZER_FORMAT, M_DEF_DIGITIZER_INIT, DigitizerIdVarPtr)
  End If
                                               
  '* allocate an image buffer
  If (Not (ImageIdVarPtr = M_NO_ALLOC) And Not (SystemIdVarPtr = M_NO_ALLOC) And Not (SystemIdVarPtr = M_NO_ALLOC)) Then
     Call LocalBufferAllocDefault(SystemIdVarPtr, DisplayIdVarPtr, DigitizerIdVarPtr, ImageIdVarPtr)
  End If

  '* enable keying if keying is supported
  If ((Not (DisplayIdVarPtr) = M_NO_ALLOC) And (Not (DisplayIdVarPtr) = M_NO_ALLOC) And (Not (M_DEF_DISPLAY_KEY_ENABLE_ON_ALLOC) = 0) And (MdispInquire(DisplayIdVarPtr, M_DISP_KEY_SUPPORTED, 0))) Then
     Call MdispOverlayKey(DisplayIdVarPtr, M_KEY_ON_COLOR, M_EQUAL, &HFF&, M_DEF_DISPLAY_KEY_COLOR)
  End If

End Sub


'/************************************************************************/
'/* MsysFreeDefault - macro to free default MIL objects:                 */
'/*                                                                      */
'/* MIL_ID ApplicationId;                                                */
'/* MIL_ID SystemId;                                                     */
'/* MIL_ID DisplayId;                                                    */
'/* MIL_ID DigitizerId;                                                  */
'/* MIL_ID ImageId;                                                      */
'/*                                                                      */
'/************************************************************************/

Sub MappFreeDefault(ApplicationId As Long, SystemId As Long, DisplayId As Long, DigitizerId As Variant, BufferId As Long)
                                                                           
  '* free the image buffer
  If Not (BufferId = M_NO_ALLOC) Then
  Call MbufFree(BufferId)
  End If
                                
  '* free digitizer
  If Not (DigitizerId = M_NO_ALLOC) Then
  Call MdigFree(DigitizerId)
  End If
                                
  '* free the display
  If (Not (DisplayId) = M_NO_ALLOC) Then

     If ((Not (M_DEF_DISPLAY_KEY_DISABLE_ON_FREE) = 0) And (MdispInquire((DisplayId), M_DISP_KEY_SUPPORTED, M_NULL) = M_YES)) Then
        Call MdispOverlayKey((DisplayId), M_KEY_OFF, M_NULL, M_NULL, M_NULL)
     End If
     
  Call MdispFree((DisplayId))
  End If
  
  '* free the system
  If Not (SystemId = M_NO_ALLOC) Then
  Call MsysFree(SystemId)
  End If
                                
  '* free the system
  If Not (ApplicationId = M_NO_ALLOC) Then
  Call MappFree(ApplicationId)
  End If
End Sub




'******************************************************************************
