using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Runtime.InteropServices;

namespace ForceDefaultPort
{
    class DeviceCfg
    {
        [DllImport("MSPorts.dll", SetLastError = true)]
        //LONG ComDBOpen(PHCOMDB PHComDB);
        static extern int ComDBOpen(out IntPtr hComDB);
        [DllImport("MSPorts.dll", SetLastError = true)]
        //LONG ComDBClose(HCOMDB HComDB);
        static extern long ComDBClose(IntPtr HCOMDB);
        [DllImport("MSPorts.dll", SetLastError = true)]
        //LONG ComDBGetCurrentPortUsage(IN HCOMDB HComDB, IN OUT PBYTE Buffer, IN DWORD BufferSize, IN ULONG ReportType, OUT LPDWORD MaxPortsReported);
        static extern long ComDBGetCurrentPortUsage(IntPtr HCOMDB, [In, Out] byte[] buffer, int bufferSize, int reportType, [Out] out int maxPortsReported);
        [DllImport("MSPorts.dll", CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        //LONG ComDBClaimPort(HCOMDB HComDB, DWORD  ComNumber, BOOL ForceClaim, PBOOL  Forced);
        static extern long ComDBClaimPort(IntPtr HCOMDB, int ComNumber, [MarshalAs(UnmanagedType.Bool)] bool ForceClaim, [MarshalAs(UnmanagedType.Bool)] out bool Force);
        [DllImport("MSPorts.dll", SetLastError = true)]
        //LONG ComDBReleasePort(HCOMDB HComDB, DWORD ComNumber);
        static extern long ComDBReleasePort(IntPtr HCOMDB, int ComNumber);
        [DllImport("MSPorts.dll", SetLastError = true)]
        //LONG ComDBClaimNextFreePort(HCOMDB HComDB, LPDWORD ComNumber);
        static extern long ComDBClaimNextFreePort(UInt32 HCOMDB, IntPtr ComNumber);

        // Rescan for hardware changes
        [DllImport("CfgMgr32.dll", SetLastError = true)]
        public static extern int CM_Locate_DevNodeA(ref int pdnDevInst, string pDeviceID, int ulFlags);
        [DllImport("CfgMgr32.dll", SetLastError = true)]
        public static extern int CM_Reenumerate_DevNode(int dnDevInst, int ulFlags);

        public const int CM_LOCATE_DEVNODE_NORMAL = 0x00000000;
        public const int CM_REENUMERATE_NORMAL = 0x00000000;
        public const int CR_SUCCESS = 0x00000000;

        public const int COMDB_MIN_PORTS_ARBITRATED = 256;
        public const int COMDB_MAX_PORTS_ARBITRATED = 4096;

        public const int CDB_REPORT_BITS = 0x00000000;
        public const int CDB_REPORT_BYTES = 0x00000001;

        /********************************************************************************************************/
        // ATTRIBUTES
        /********************************************************************************************************/
        #region -- attributes --

        const string INGNAR = "0b00"; //Do NOT make this uppercase
        const string IDTECH = "0acd";
        const string VERIFN = "11ca";
        const string IdTechString = "idtech";

        List<int> comPorts = new List<int>(new int[] { 30, 31, 32, 33, 34, 35, 109, 110, 111, 112, 113 });
        int targetPort;

        #endregion

        public void DeviceInit(int first, int last)
        {
            if (first > 0 && last > 0)
            {
                ClearUpCommPorts(first, last);
            }
            else
            {
                //string result = ReportUSBCommPorts().TrimStart(new char[] { 'C', 'O', 'M' });
                //targetPort = Convert.ToInt32(result);
                //if (targetPort > 0)
                //{
                //    comPorts.Add(targetPort);
                //}
                ReportAllCommPorts();
            }
        }

        public void ListPortsInUse(int start, int end)
        {
            ListInUsePorts(start, end);
        }

        public void SetPortsInUse(int start, int end)
        {
            SetInUsePorts(start, end);
        }

        public bool FindIngenicoDevice(ref string description, ref string deviceID)
        {
            List<USBDeviceInfo> devices = GetUSBDevices();
            if (devices.Count == 1)
            {
                BoolStringDuple output = output = new BoolStringDuple(true, devices[0].Description.ToLower().Replace("ingenico ", ""));
                deviceID = devices[0].DeviceID;
                description = devices[0].Description;

                return true;
            }
            return false;
        }

        public bool FindVerifoneDevice(ref string description, ref string deviceID)
        {
            List<USBDeviceInfo> devices = GetUSBDevices();
            if (devices.Count >= 1)
            {
                BoolStringDuple output = output = new BoolStringDuple(true, devices[0].Description.ToLower().Replace("verifone ", ""));
                deviceID = devices[0].DeviceID;
                description = devices[0].Description;

                return true;
            }
            return false;
        }

        public static List<USBDeviceInfo> GetUSBDevices()
        {
            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();
            ManagementObjectCollection collection;
            try
            {
                using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_PnPEntity"))
                {
                    collection = searcher.Get();
                }
                foreach (var device in collection)
                {
                    var deviceID = ((string)device.GetPropertyValue("DeviceID") ?? "").ToLower();
                    if (string.IsNullOrWhiteSpace(deviceID))
                    {
                        continue;
                    }
                    Debug.WriteLine($"device: {deviceID}");
                    if (deviceID.Contains("usb\\") && (deviceID.Contains($"vid_{INGNAR}") || deviceID.Contains($"vid_{IDTECH}") || deviceID.Contains($"vid_{VERIFN}")))
                    {
                        devices.Add(new USBDeviceInfo(
                            (string)device.GetPropertyValue("DeviceID"),
                            (string)device.GetPropertyValue("PNPDeviceID"),
                            (deviceID.Contains($"vid_{IDTECH}") ? DeviceCfg.IdTechString : (string)device.GetPropertyValue("Description"))
                        ));
                    }
                }
                collection.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return devices;
        }

        private string ReportUSBCommPorts()
        {
            string port = string.Empty;

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\WMI", "SELECT * FROM MSSerial_PortName");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    //If the serial port's instance name contains USB 
                    //it must be a USB to serial device
                    if (queryObj["InstanceName"].ToString().Contains("USB"))
                    {
                        Console.WriteLine("----------------------------------------------------------------------");
                        Console.WriteLine("MSSerial_PortName instance");
                        Console.WriteLine("----------------------------------------------------------------------");
                        Console.WriteLine("InstanceName : {0}", queryObj["InstanceName"]);
                        Console.WriteLine("Configuration: " + queryObj["PortName"] + " - (USB to SERIAL adapter/converter)");
                        port = queryObj["PortName"].ToString();
                        //SerialPort p = new SerialPort(port);
                        //p.PortName = "COM11";
                        //return port;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return port;
        }

        private string ReportAllCommPorts()
        {
            string portsArray = string.Empty;

            try
            {
                //ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\WMI", "SELECT * FROM MSSerial_PortName");

                //Console.WriteLine("----------------------------------------------------------------------");
                //Console.WriteLine("REPORT ALL SERIAL COMM PORTS");
                //Console.WriteLine("----------------------------------------------------------------------");

                //foreach (ManagementObject queryObj in searcher.Get())
                //{
                //    Console.WriteLine("InstanceName : {0}", queryObj["InstanceName"]);
                //    Console.WriteLine("Configuration: " + queryObj["PortName"] + " - (USB to SERIAL adapter/converter)");
                //    port = queryObj["PortName"].ToString();
                //}
                // Get a list of serial port names.
                string[] ports = SerialPort.GetPortNames();
                foreach (string port in ports)
                {
                    Console.WriteLine(port);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return portsArray;
        }

        private void SetInUsePorts(int start, int end)
        {
            int state = ComDBOpen(out IntPtr PHCOMDB);
            if (PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
            {
                Console.WriteLine("SET PORT(S) IN USE ------------------------------------------------------------------------\r\n");

                for (int port = start; port <= end; port++)
                {
                    long result = ComDBClaimPort(PHCOMDB, port, true, out bool forced);
                    Console.WriteLine($"device: COM{port} forced in-use with status={result}.");
                }
                long dsfdf1 = ComDBClose(PHCOMDB);
            }
        }

        private void ListInUsePorts(int start, int end)
        {
            int state = ComDBOpen(out IntPtr PHCOMDB);
            if (PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
            {
                int maxPortsReported = 0;
                byte[] buffer = null;
                long result = ComDBGetCurrentPortUsage(PHCOMDB, buffer, 0, CDB_REPORT_BYTES, out maxPortsReported);
                if (result == (int)ERROR_STATUS.ERROR_SUCCESS)
                {
                    buffer = new byte[maxPortsReported];
                    result = ComDBGetCurrentPortUsage(PHCOMDB, buffer, maxPortsReported, CDB_REPORT_BYTES, out maxPortsReported);
                    if (result == (int)ERROR_STATUS.ERROR_SUCCESS)
                    {
                        Console.WriteLine("LIST PORT(S) IN USE ----------------------------------------------------------------------\r\n");
                        // port number is an integer that ranges from 1 to COMDB_MAX_PORTS_ARBITRATED
                        for (int port = start - 1; port < end; port++)
                        {
                            if (buffer[port] == 0x01)
                            {
                                Console.WriteLine($"COM{port + 1}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"ComDBGetCurrentPortUsage: error={result}");
                }
                long dsfdf1 = ComDBClose(PHCOMDB);
            }
        }

        private void ClearUpCommPorts(int first, int last)
        {
            try
            {
                // Clear in-use COM Port
                int state = ComDBOpen(out IntPtr PHCOMDB);
                if (PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
                {
                    Console.WriteLine("");
                    Console.WriteLine("CLEAR PORT(S) IN USE ---------------------------------------------------------------------\r\n");
                    for (int index = first; index <= last; index++)
                    {
                        long status = ComDBReleasePort(PHCOMDB, index);
                        Console.WriteLine($"device: COM{index} released with status={status}.");
                    }
                    long dsfdf1 = ComDBClose(PHCOMDB);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void RescanForHardwareChanges()
        {
            int pdnDevInst = 0;

            if (CM_Locate_DevNodeA(ref pdnDevInst, null, CM_LOCATE_DEVNODE_NORMAL) != CR_SUCCESS)
            {
                Console.WriteLine("device: failed to locate hardware devices");
            }
            if (CM_Reenumerate_DevNode(pdnDevInst, CM_REENUMERATE_NORMAL) != CR_SUCCESS)
            {
                Console.WriteLine("Failed to reenumerate hardware devices");
            }
            else
            {
                Console.WriteLine("device: reenumerated hardware devices");
            }
        }
    }

    enum ERROR_STATUS
    {
        ERROR_SUCCESS = 0,
        ERROR_ACCESS_DENIED = 5,
        ERROR_INVALID_PARAMETER = 87,
        ERROR_NOT_CONNECTED = 2250
    };
}
