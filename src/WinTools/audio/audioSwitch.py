import utils.policyconfig as pc
import comtypes

def switchSpeaker(switchTo, role):
    policy_config = comtypes.CoCreateInstance(
        pc.CLSID_PolicyConfigClient,
        pc.IPolicyConfig,
        comtypes.CLSCTX_ALL
    )
    pc.ERole.eCommunications
    policy_config.SetDefaultEndpoint(switchTo, role)