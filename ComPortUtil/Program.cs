using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ForceDefaultPort
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"\r\n==========================================================================================");
            Console.WriteLine($"{Assembly.GetEntryAssembly().GetName().Name} - Version {Assembly.GetEntryAssembly().GetName().Version}");
            Console.WriteLine($"==========================================================================================\r\n");

            if (args.Length > 0)
            {
                switch (args[0].ToUpper())
                {
                    case "/INFO":
                    {
                        try
                        {
                            DeviceCfg config = new DeviceCfg();
                            if (config != null)
                            {
                                config.DeviceInit(-1, -1);

                                //string description = string.Empty;
                                //string deviceID = string.Empty;
                                //config.FindVerifoneDevice(ref description, ref deviceID);
                                //Console.WriteLine($"verifone: ID={deviceID}, DESC={description}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"main: exception={ex.Message}");
                        }
                        break;
                    }
                    case "/CLEAR":
                    {
                        if (args.Length == 3)
                        {
                            int.TryParse(args[1], out int first);
                            int.TryParse(args[2], out int last);
                            if (first > 0 && last > 0 && (first <= last))
                            { 
                                try
                                {
                                    DeviceCfg config = new DeviceCfg();
                                    if (config != null)
                                    {
                                        config.DeviceInit(first, last);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"main: exception={ex.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Verb [/CLEAR]: Invalid Port Parameter(s) '{args[1]} | {args[2]}'");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Verb [/CLEAR]: Missing Parameters - [PORT1 PORTN]");
                        }
                        break;
                    }
                    case "/SET":
                    {
                        if (args.Length == 3)
                        {
                            int.TryParse(args[1], out int first);
                            int.TryParse(args[2], out int last);
                            if (first > 0 && last > 0 && (first <= last))
                            {
                                try
                                {
                                    DeviceCfg config = new DeviceCfg();
                                    if (config != null)
                                    {
                                        config.SetPortsInUse(first, last);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"main: exception={ex.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Verb [/SET]: Invalid Port Parameter(s) '{args[1]} | {args[2]}'");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Verb [/SET]: Missing Parameters - [PORT1 PORTN]");
                        }
                        break;
                    }
                    case "/LIST":
                        {
                            if (args.Length == 3)
                            {
                                int.TryParse(args[1], out int first);
                                int.TryParse(args[2], out int last);
                                if (first > 0 && last > 0 && (first <= last))
                                {
                                    try
                                    {
                                        DeviceCfg config = new DeviceCfg();
                                        if (config != null)
                                        {
                                            config.ListPortsInUse(first, last);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"main: exception={ex.Message}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"Verb [/LIST]: Invalid Port Parameter(s) '{args[1]} | {args[2]}'");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Verb [/LIST]: Missing Parameters - [PORT1 PORTN]");
                            }
                            break;
                        }
                }
            }
            else
            {
                Console.WriteLine($"Missing Verb - [/INFO] [/LIST] [/SET] [/CLEAR]");
            }
        }
    }
}
