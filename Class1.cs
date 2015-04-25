using System;
using System.IO;
using System.Management;
using System.Runtime.Serialization.Formatters.Binary;
using MSI.Afterburner;
using MSI.Afterburner.Exceptions;
using LcdStudio.CoreInterfaces;

namespace MSIAfterburnerPlugin
{
    public class Plugin : AbstractDataPlugin
    {
        public override void RegisterData(IDataService ds)
        {
            ds.RegisterVariable(Group, "JamiePlugin.CPUSpeed", "JamiePlugin\\CPUSpeed");
            ds.RegisterVariable(Group, "JamiePlugin.Today", "JamiePlugin\\Today");
            ds.RegisterVariable(Group, "JamiePlugin.TotalMem", "JamiePlugin\\TotalMem");
            ds.RegisterVariable(Group, "JamiePlugin.AfterburnerCoreClock", "JamiePlugin\\AfterburnerCoreClock");
            ds.RegisterVariable(Group, "JamiePlugin.AfterburnerMemClock", "JamiePlugin\\AfterburnerMemClock");
            ds.RegisterVariable(Group, "JamiePlugin.AfterburnerShaderClock", "JamiePlugin\\AfterburnerShaderClock");
            ds.RegisterVariable(Group, "JamiePlugin.AfterburnerFanSpeed", "JamiePlugin\\AfterburnerFanSpeed");
            this.UpdateInterval = 1000; //Update every 1000ms (1 sec) 
        }

        public override void UpdateData(IDataService ds)
        {
            ds.SetValue("JamiePlugin.Today", DayOfWeek());
            ds.SetValue("JamiePlugin.CPUSpeed", CPUSpeed());
            ds.SetValue("JamiePlugin.TotalMem", DisplayTotalRam());
            ds.SetValue("JamiePlugin.AfterburnerCoreClock", MSIConnect(0));
            ds.SetValue("JamiePlugin.AfterburnerMemClock", MSIConnect(1));
            ds.SetValue("JamiePlugin.AfterburnerShaderClock", MSIConnect(2));
            ds.SetValue("JamiePlugin.AfterburnerFanSpeed", MSIConnect(3));
        }

        public string CPUSpeed()
        {
            ManagementObject Mo = new ManagementObject("Win32_Processor.DeviceID='CPU0'");
            uint sp = (uint)(Mo["CurrentClockSpeed"]);
            Mo.Dispose();
            decimal speed = Convert.ToDecimal(sp);
            speed = Math.Round(speed / 1000, 1);
            string speedString = speed + " Ghz";
            return speedString;
        }

        public string DayOfWeek()
        {
            return System.DateTime.Now.DayOfWeek.ToString();
        }

        private static decimal DisplayTotalRam()
        {
            string Query = "SELECT MaxCapacity FROM Win32_PhysicalMemoryArray";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(Query);
            decimal finalSizeInGB = 0;
            foreach (ManagementObject WniPART in searcher.Get())
            {
                UInt32 SizeinKB = Convert.ToUInt32(WniPART.Properties["MaxCapacity"].Value);
                UInt32 SizeinMB = SizeinKB / 1024;
                UInt32 SizeinGB = SizeinMB / 1024;
                finalSizeInGB = SizeinMB / 1024;
            }
            return finalSizeInGB;
        }

        private static int MSIConnect(int inputNumber) // 0 = coreClock, 1 = memClock, 2= shaderClock, 3 = fanSpeed
        {
            ControlMemory macm = new ControlMemory();
            int coreClock = Convert.ToInt32(macm.GpuEntries[0].CoreClockCur);
            int memClock = Convert.ToInt32(macm.GpuEntries[0].MemoryClockCur);
            int shaderClock = Convert.ToInt32(macm.GpuEntries[0].ShaderClockCur);
            int fanSpeed = Convert.ToInt32(macm.GpuEntries[0].FanSpeedCur);
            if (inputNumber == 1)
            {
                return coreClock;
            }
            else if (inputNumber == 2)
            {
                return memClock;
            }
            else if (inputNumber == 3)
            {
                return shaderClock;
            }
            else if (inputNumber == 4)
            {
                return fanSpeed;
            }
            return 0;
        }
    }
}
