## This module is in large chunks taken from https://github.com/kdschlosser/pyWinCoreAudio/blob/431d5d9b470083a6a5c51738e08b12009fa315eb/pyWinCoreAudio/__core_audio/policyconfig.py.
## This requires us to use comtypes alongside pywin32, which is somewhat of a duplication in depedencies.
## I think we could use pyWinCoreAudio instead of soundcard, but we will still likely need pywin32 for the service.


import comtypes, enum
from comtypes import COMMETHOD, GUID
import ctypes
from ctypes import (
    POINTER,
    HRESULT,
     c_int as enum
)



from ctypes.wintypes import (
    INT,
    BOOL,
    LPCWSTR,
    WORD
)

IID_IPolicyConfig = GUID(
    '{f8679f50-850a-41cf-9c72-430f290290c8}'
)

CLSID_PolicyConfigClient = GUID(
    '{870af99c-171d-4f9e-af0d-e63df40c2bc9}'
)

IID_AudioSes = (
    '{00000000-0000-0000-0000-000000000000}'
)

ConfigFactory = GUID("2a59116d-6c4f-45e0-a74f-707e3fef9258")


REFERENCE_TIME = ctypes.c_longlong


LPCGUID = POINTER(GUID)
LPREFERENCE_TIME = POINTER(REFERENCE_TIME)

class DeviceSharedMode(ctypes.Structure):
    _fields_ = [
        ('dummy_', INT)
    ]


PDeviceSharedMode = POINTER(DeviceSharedMode)

class WAVEFORMATEX(ctypes.Structure):
    _fields_ = [
        ('wFormatTag', WORD),
        ('nChannels', WORD),
        ('nSamplesPerSec', WORD),
        ('nAvgBytesPerSec', WORD),
        ('nBlockAlign', WORD),
        ('wBitsPerSample', WORD),
        ('cbSize', WORD),
    ]


PWAVEFORMATEX = POINTER(WAVEFORMATEX)

class _tagpropertykey(ctypes.Structure):
    pass


class tag_inner_PROPVARIANT(ctypes.Structure):
    pass

PROPVARIANT = tag_inner_PROPVARIANT
PPROPVARIANT = POINTER(PROPVARIANT)

PROPERTYKEY = _tagpropertykey
PPROPERTYKEY = POINTER(_tagpropertykey)

class ERole(enum):
    eConsole = 0
    eMultimedia = 1
    eCommunications = 2
    ERole_enum_count = 3


class IPolicyConfig(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPolicyConfig
    _methods_ = (
        COMMETHOD(
            [],
            HRESULT,
            'GetMixFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], POINTER(PWAVEFORMATEX), 'pFormat')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'ResetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDeviceFormat',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PWAVEFORMATEX, 'pEndpointFormat'),
            (['in'], PWAVEFORMATEX, 'pMixFormat')
        ),

        COMMETHOD(
            [],
            HRESULT,
            'GetProcessingPeriod',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bDefault'),
            (['out'], LPREFERENCE_TIME, 'hnsDefaultDevicePeriod'),
            (['out'], LPREFERENCE_TIME, 'hnsMinimumDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetProcessingPeriod',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], LPREFERENCE_TIME, 'hnsDevicePeriod')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetShareMode',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['out'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetShareMode',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PDeviceSharedMode, 'pMode')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'GetPropertyValue',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['out'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetPropertyValue',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], PPROPERTYKEY, 'key'),
            (['in'], PPROPVARIANT, 'pValue')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetDefaultEndpoint',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], ERole, 'ERole')
        ),
        COMMETHOD(
            [],
            HRESULT,
            'SetEndpointVisibility',
            (['in'], LPCWSTR, 'pwstrDeviceId'),
            (['in'], BOOL, 'bVisible')
        )
    )


PIPolicyConfig = POINTER(IPolicyConfig)

class AudioSes(object):
    name = u'AudioSes'
    _reg_typelib_ = (IID_AudioSes, 1, 0)


class CPolicyConfigClient (comtypes.CoClass):
    _reg_clsid_ = CLSID_PolicyConfigClient
    _idlflags_ = []
    _reg_typelib_ = (IID_AudioSes, 1, 0)
    _com_interfaces_ = [IPolicyConfig]
