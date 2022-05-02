using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Text.Json;
using Newtonsoft.Json;

namespace ExPlus
{

    public class BackEnd
    {


        public static Tuple<IntPtr, IntPtr> HandleProc()
        {
            var hndls = Macros.getHandles();
            IntPtr? hwn = null, htree = null;
            while (
                   hwn == null || (int)hwn == 0 
            || htree == null || (int)htree == 0
                )
            {
                 hndls = Macros.getHandles();
                 hwn = hndls.Item1;
                 htree = hndls.Item2;
                Thread.Sleep(500);
            }
            return hndls;

        }
        static void Main(string[] args)
        {


            var hs = HandleProc();
            IntPtr hwn = hs.Item1, htree = hs.Item2;

            int num = -1;
            while (true)
            {
                var tnum = Macros.GetSelectedInstrumentNum(hwn, htree);
                //if (tnum != -1 && tnum != num)
                if (tnum == 1)
                {
                    Console.WriteLine(tnum);
                    Macros.SelectAllVoices(hwn, htree, tnum);
                    var dt = Macros.GetVoiceProcessingDatas(hwn, htree);
                    var json = JsonConvert.SerializeObject(dt);
                    //Console.WriteLine(json);
                    //Macros.VoiceProcessingToClipboard(hwn);

                    num = tnum;
                }
                Thread.Sleep(10000);


            }


        }
    }
}
