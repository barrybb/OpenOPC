###########################################################################
#
# OpenOPC Gateway Service
#
# A Windows service providing remote access to the OpenOPC library.
#
# Copyright (c) 2007-2015 Barry Barnreiter (barrybb@gmail.com)
#
###########################################################################

import win32serviceutil
import win32service
import win32event
import servicemanager
import winerror
import _winreg
import select
import socket
import os
import sys
import time
import threading
import OpenOPC

try:
    import Pyro.core
    import Pyro.protocol
    import Pyro.errors
except ImportError:
    print 'Pyro module required (http://pyro.sourceforge.net/)'
    exit()

Pyro.config.PYRO_MULTITHREADED = 1

opc_class = OpenOPC.OPC_CLASS
opc_gate_host = None
opc_gate_port = 7766
inactive_timeout = 60
max_clients = 25

def getvar(env_var):
    """Read system enviornment variable from registry"""
    try:
        key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'SYSTEM\\CurrentControlSet\\Control\Session Manager\Environment',0,_winreg.KEY_READ)
        value, valuetype = _winreg.QueryValueEx(key, env_var)
        return value
    except:
        return None

# Get env vars directly from the Registry since a reboot is normally required
# for the Local System account to inherit these.

if getvar('OPC_CLASS'):  opc_class = getvar('OPC_CLASS')

if getvar('OPC_GATE_HOST'):
    opc_gate_host = getvar('OPC_GATE_HOST')
    if opc_gate_host.strip() in ('0.0.0.0', ''):
        opc_gate_host = None

if getvar('OPC_GATE_PORT'):  opc_gate_port = int(getvar('OPC_GATE_PORT'))
if getvar('OPC_MAX_CLIENTS'): max_clients = int(getvar('OPC_MAX_CLIENTS'))
if getvar('OPC_INACTIVE_TIMEOUT'): inactive_timeout = int(getvar('OPC_INACTIVE_TIMEOUT'))

def inactive_cleanup(exit_event):
    while True:
        exit_event.wait(60.0)

        if exit_event.is_set():
            all_sessions = OpenOPC.get_sessions(host=opc_gate_host, port=opc_gate_port)
            for oid, ip, ctime, xtime in all_sessions:
                OpenOPC.close_session(oid, host=opc_gate_host, port=opc_gate_port)
            exit_event.clear()
            return
        else:
            try:
                all_sessions = OpenOPC.get_sessions(host=opc_gate_host, port=opc_gate_port)
            except Pyro.errors.ProtocolError:
                all_sessions = []
            if len(all_sessions) > max_clients:
                stale_sessions = sorted(all_sessions, key=lambda s: s[3])[:-max_clients]
            else:
                stale_sessions = [s for s in all_sessions if time.time()-s[3] > (inactive_timeout*60)]
            for oid, ip, ctime, xtime in stale_sessions:
                OpenOPC.close_session(oid, host=opc_gate_host, port=opc_gate_port)                
                time.sleep(1)

class opc_client(Pyro.core.ObjBase, OpenOPC.client):
    def __init__(self, remote_ip=''):
        Pyro.core.ObjBase.__init__(self)
        OpenOPC.client.__init__(self)
        self.remote_ip = remote_ip

class opc(Pyro.core.ObjBase):
    def __init__(self):
        Pyro.core.ObjBase.__init__(self)
        self._remote_hosts = {}
        self._opc_objects = {}
        self._init_times = {}

    def get_clients(self):
        """Return list of server instances as a hash of GUID:host"""
        
        reg = self.getDaemon().getRegistered()
        hosts = self._remote_hosts
        init_times = self._init_times
        objects = self._opc_objects
        
        hlist = [(k, hosts[k] if hosts.has_key(k) else '', init_times[k], objects[k].lastUsed) for k,v in reg.iteritems() if v == None]
        return hlist

    def force_close(self, guid):
        obj = self._opc_objects[guid]
        servicemanager.LogInfoMsg('\n\nForcing closed session started by %s' % obj.remote_ip)
        obj.close()
    
    def create_client(self):
        """Create a new OpenOPC client instance in the Pyro server"""

        remote_ip = self.getLocalStorage().caller.addr[0]
        servicemanager.LogInfoMsg('\n\nConnect from %s' % remote_ip)

        opc_obj = opc_client(remote_ip)
        uri = self.getDaemon().connect(opc_obj)
        
        opc_obj._open_serv = self
        opc_obj._open_self = opc_obj
        opc_obj._open_host = self.getDaemon().hostname
        opc_obj._open_port = self.getDaemon().port        
        opc_obj._open_guid = uri.objectID

        self._remote_hosts[opc_obj.GUID()] = '%s' % (remote_ip)
        self._opc_objects[opc_obj.GUID()] = opc_obj
        self._init_times[opc_obj.GUID()] =  time.time()

        return Pyro.core.getProxyForURI(uri)

    def release_client(self, obj):
        """Release a OpenOPC client instance in the Pyro server"""

        remote_ip = self.getLocalStorage().caller.addr[0]
        servicemanager.LogInfoMsg('\n\nDisconnect from %s' % remote_ip)

        self.getDaemon().disconnect(obj)
        del self._remote_hosts[obj.GUID()]
        del self._opc_objects[obj.GUID()]
        del self._init_times[obj.GUID()]
        del obj
   
class OpcService(win32serviceutil.ServiceFramework):
    _svc_name_ = "zzzOpenOPCService"
    _svc_display_name_ = "OpenOPC Gateway Service"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
    
    def SvcStop(self):
        servicemanager.LogInfoMsg('\n\nStopping service')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg('\n\nStarting service\nOPC_GATE_HOST=%s\nOPC_GATE_PORT=%d\nOPC_MAX_CLIENTS=%d\nOPC_INACTIVE_TIMEOUT=%d' % ('0.0.0.0' if opc_gate_host is None else opc_gate_host, opc_gate_port, max_clients, inactive_timeout))

        exit_event = threading.Event()
        p = threading.Thread(target=inactive_cleanup, args=(exit_event,))
        p.start()

        daemon = Pyro.core.Daemon(host=opc_gate_host, port=opc_gate_port)
        daemon.connect(opc(), "opc")

        stop_pending = False

        while True:
            if not stop_pending and win32event.WaitForSingleObject(self.hWaitStop, 0) == win32event.WAIT_OBJECT_0:
                exit_event.set()
                stop_pending = True
            elif stop_pending and not exit_event.is_set():
                break
    
            socks = daemon.getServerSockets()
            ins,outs,exs = select.select(socks,[],[],1)
            for s in socks:
                if s in ins:
                    daemon.handleRequests()
                    break

        p.join()
        daemon.shutdown()
        
if __name__ == '__main__':
    if len(sys.argv) == 1:
        try:
            evtsrc_dll = os.path.abspath(servicemanager.__file__)
            servicemanager.PrepareToHostSingle(OpcService)
            servicemanager.Initialize('zzzOpenOPCService', evtsrc_dll)
            servicemanager.StartServiceCtrlDispatcher()
        except win32service.error, details:
            if details[0] == winerror.ERROR_FAILED_SERVICE_CONTROLLER_CONNECT:
                win32serviceutil.usage()
    else:
        win32serviceutil.HandleCommandLine(OpcService)
