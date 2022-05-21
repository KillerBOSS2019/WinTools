from threading import Thread

""" Import TouchPortalAPI stuff """
from constants import TYPES, TPClient

""" Import callback methods """
from callback.NotificationClicked import NotificationClicked
from callback.onAction import onAction
from callback.onConnect import onConnect
from callback.onConnectorChange import onConnectorChange
from callback.onHoldDown import onHoldDown
from callback.onListChange import onListChange
from stateUpdate import stateUpdate

""" Run onConnect function when TP Connects to TP """
TPClient.on(TYPES.onConnect, onConnect)

"""  Run onAction function when TP sends action event """
TPClient.on(TYPES.onAction, onAction)

"""  Run onListChange when user selects item form action drop-down menu """
TPClient.on(TYPES.onListChange, onListChange)

""" Run onHoldDown when user pressing down a button """
TPClient.on(TYPES.onHold_down, onHoldDown)

""" Run onConnectorChange when user slide a slider """
TPClient.on(TYPES.onConnectorChange, onConnectorChange)

if __name__ == "__main__":
    Thread(target=stateUpdate).start()
    try:
        TPClient.connect()
    except KeyboardInterrupt:
        pass
    except Exception:
        from traceback import format_exc
        print(f"Exception in TP Client:\n{format_exc()}")
    finally:
        TPClient.disconnect()
