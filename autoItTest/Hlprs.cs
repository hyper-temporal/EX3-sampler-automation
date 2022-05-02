using System;
using System.Collections.Generic;
using System.Text;
using AutoIt;
using System.Text.Json;
using System.Globalization;
using System.Threading;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.Windows;

namespace ExPlus
{

    #region EX Structs
    public struct LinkData
    {
        public int? presetNumber;

        public float? volume;
        public int? pan;
        public int? xpoz;
        public float? fine;

        public ControlRouting routing;
    }

    public struct LowHighPlacement
    {
        public string low;
        public int? fadel;
        public string high;
        public int? fadeh;
    }

    public struct SourcePlacement
    {
        public string source;
        public LowHighPlacement placement;
    }

    public struct Detail
    {
        public bool mute;
        public bool solo;
        public int? groop;
        public string sample;       //code for if it's a zone 
        public string direction;
        public LowHighPlacement notePlacement;
        public string orK;
        public float? volume;
        public int? pan;
        public int? xpoz;
        public int? coarse;
        public float? fine;            
    }
    public struct ControlRouting
    {
        public LowHighPlacement Key;
        public LowHighPlacement Vel;
        public LowHighPlacement RT;
        public SourcePlacement CC1, CC2, CC3, CC4, CC5;
    }

    
    public struct Zone
    {
        public Detail detail;
        public ControlRouting routing;
    };


    public struct Tune
    {
        public string
         transpose
       , Coarse
       , Fine
       ;
    }
    public struct Chorus
    {
        public string
           amount
       , width
       , initms
       ;
    }
    public struct Twistaloop
    {
        public string
           SpeedPrcnt
       , SoopPrcnt
       ;
        public bool StartAtLoop;
    }
    public struct VPKey
    {
        public string delay;
        public float offset;
        public bool latch;
        public bool bpm;
        public string assignGroup;
        public string keymode;
    }
    public struct Bend
    {
        public string Up;
        public string Dn;
    }
    public struct Glide
    {
        public string Curve;
        public float Rate;
    }
    public struct Env
    {
        public bool bpm;
        public string Attack1Time
            , Attack1Level
            , Attack2Time
            , Attack2Level
            , Decay1Time
            , Decay1Level
            , Decay2Time
            , Decay2Level
            , Release1Time
            , Release1Level
            , Release2Time
            , Release2Level
            ;
    }
    public struct LFO
    {
        public bool keySync, bpm;
        public string Frequency
            , Delay
            , Variation
            ;
    }
    public class Step
    {
        public float value; //-32. to 32.0
        public bool trigger;
        public override string ToString()
        {
            return $"v={value} t={trigger}";
        }
    }
    public class FuncGen
    {
        public string StepRate;
        public bool bpm;
        public string sync;
        public bool smooth;
        public string direction;
        public int endStep;


        public Step[] steps; //64

    }
    public class Cord
    {
        public string source, destination;
        public float value;
    }
    public class MorphParameters
    {
        public float lofreq, hifreq, gain;
    }
    public class FilterDesignerStage
    {
        public string ftype;
        public float freq1, gain1;
        public float freq2, gain2;
    }

    public class FilterMorpher
    {
        public MorphParameters left, right;
    }

    public class FilterDesigner
    {
        public FilterDesignerStage[] stages;
    }





    public class Filter
    {
        public string name;
        public float freq, rez;
        public FilterMorpher morpher;
        public FilterDesigner designer;

    }
    public class Amplifier
    {
        public float Volume, Pan, FxWetDry, AmpEnvRange;
        public int Main, Aux1, Aux2, Aux3;
        public bool ClassicResponse;
    }
    public class VoiceProcessing
    {
        public Tune tune { get; set; }
        public Chorus chorus { get; set; }
        public Twistaloop twistaloop { get; set; }
        public Bend bend { get; set; }
        public Glide glide { get; set; }
        public VPKey key { get; set; }
        public Env AmpEnv { get; set; }
        public Env FilterEnv { get; set; }
        public Env AuxEnv { get; set; }
        public Filter filtre { get; set; }
        public Amplifier amp { get; set; }

        public LFO lfo1 { get; set; }
        public LFO lfo2 { get; set; }
        public FuncGen fg1, fg2, fg3;
        public Cord[] cords;
    };

   

    #endregion

    public static class Macros
    {


        public static void SetSelectionSamplesParamsPt1(
           IntPtr hw,
           bool mute,
           bool solo,
           int volume,
           int pan,
           float fineTune,
           int coarseTune,
           int transpose
           )
        {
            SetBooleanCtrlValue(hw, 111, mute);
            SetBooleanCtrlValue(hw, 112, solo);
            SetEditValue(hw, 258, ((float)volume).ToString());
            SetEditValue(hw, 259, pan.ToString());
            SetEditValue(hw, 260, ((float)fineTune).ToString());
            SetEditValue(hw, 261, coarseTune.ToString());
            SetEditValue(hw, 262, transpose.ToString());
        }


       
        static Tuple<int, IntPtr> GetNextLinkPresetNum(
            IntPtr hw,
            IntPtr ht,
            int pst
        ){
            //var hh = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(1));
            //AutoItX.ControlClick(hw, hh, "left", 3);
            //AutoItX.ControlSend(hw, hh, "{ESC}");
            if (selectNthPresetVoiceLinks(hw, ht, pst) == "1")
            {
                var hc = AutoItX.ControlGetHandle(hw, ClassName_cbb_text(5));
                var txt = GetComboboxValue(hw, 5); ;
                if (txt != String.Empty)
                {
                    txt = txt.Remove(0, 1);
                    txt = txt.Remove(4, txt.Length - 4);
                    return new Tuple<int, IntPtr>(int.Parse(txt), hc);
                }

            }
            return new Tuple<int, IntPtr>(-1, new IntPtr(0));
        }

        public static bool ForceToUpdateContext(
            IntPtr hw,
            float x,
            float y
            ){
            var re = EnableWindow(hw, false);
            var rwt = AutoItX.WinSetOnTop(hw, 1);
            WindowsUIF.RECT rc;
            var rwr = WindowsUIF.GetWindowRect(hw, out rc);
            MouseOperations.SetCursorPosition((int)(rc.X + x * rc.Width), (int)(rc.Y + y * rc.Height));
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);

            //var rmw = Screenshooter.MoveWindow(hw, 0, 0, rc.Width, rc.Height, false);
            //var rws = AutoItX.WinSetState(
            //    hw,
            //    AutoItX.SW_SHOWMAXIMIZED
            //    );
            UnprotectWindow(hw);
            return true;
        }

        public static Tuple<int, IntPtr> CutFirstLink(
            IntPtr hw,
            IntPtr ht,
            int dupnum)
        {
            var ret = GetNextLinkPresetNum(hw, ht, dupnum);
            if (ret.Item1 >= 0)
            {
                var han = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(1));
                if ((int)han == 0) return new Tuple<int, IntPtr>(-1, new IntPtr(0));
                var r = ForceToUpdateContext(hw, (float)299 / 957, (float)180 / 716);
                AutoItX.ControlClick(hw, han, "right");
                AutoItX.ControlSend(hw, han, "{DOWN 2}{ENTER}");
            }

            return ret;
        }

        public static void PresetLinksToNewPreset(
            IntPtr hw,
            IntPtr ht,
            int pst
            )
        {
            var dupnum = DuplicatePresets(hw, ht, pst);
            var t = true;
            while (t)
            {
                selectNthPresetVoiceLinks(hw, ht, dupnum);
                var l = Get1stLinkData(hw);
                if(l.presetNumber == null)
                {
                    t = false;
                }
                else
                {
                    var ltgt = CutFirstLink(hw, ht, dupnum);
                    t = ltgt.Item1 >= 0;
                    if (t)
                    {
                        SelectAllVoices(hw, ht, ltgt.Item1);
                        selectNthPresetVoicsAndZone(hw, ht, ltgt.Item1);
                        var gh = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(1));
                        if ((int)gh > 0)
                        {
                            //copy all voices
                            AutoItX.ControlClick(hw, gh, "right");
                            Thread.Sleep(200);
                            AutoItX.ControlSend(hw, gh, "{DOWN}{RIGHT}{DOWN 2}{ENTER}");
                            //once voices are copied I access their global routing ie "Win" structure
                            //And I import data from the extracted link
                            //IF it's default ( not used ) !
                            //if it's not ( not default or empty) then I copy in next available ccwin

                            SelectAllVoices(hw, ht, dupnum);
                            selectNthPresetVoicsAndZone(hw, ht, dupnum);

                            //paste voices
                            gh = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(1));
                            if ((int)gh > 0)
                            {
                                bool rr = false;
                                while (!rr)
                                {
                                    AutoItX.ControlClick(hw, gh, "right");
                                    Thread.Sleep(200);
                                    AutoItX.ControlSend(hw, gh, "{DOWN}{RIGHT}{DOWN 3}{ENTER}");
                                    int dupPos = -1;
                                    Thread.Sleep(200);
                                    var popwh = AutoItX.WinGetHandle("Paste");
                                    rr = !((int)popwh == 0);
                                    Thread.Sleep(200);
                                    if (rr)
                                    {
                                        var rbh = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:3]");
                                        var selh = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:6]");
                                        var okbutt = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:7]");
                                        if (!((int)rbh > 0 && (int)selh > 0 && (int)okbutt > 0))
                                        {
                                            throw new Exception("PasteDialog: buttons no found");
                                        }
                                        var posCtrlH = AutoItX.ControlGetHandle(popwh, GetEditClassName(1));
                                        if (!((int)posCtrlH > 0))
                                        {
                                            throw new Exception("PasteDialog: target control not found");
                                        }
                                        dupPos = int.Parse(GetEditValue(popwh, 1));

                                        AutoItX.ControlClick(popwh, rbh);   //dest will be at suggested prese 
                                        AutoItX.ControlClick(popwh, selh); //deselect move preset option
                                        AutoItX.ControlClick(popwh, okbutt);
                                    }
                                }


                            }
                            var z = getVPSampleZone(hw);
                            CopyLinkDataToVP(hw, ht, ref l, ref z, dupnum);

                            //Thread.Sleep(1000);
                        }
                    }

                }
               
            }
            //Macros.DeletePreset(hw, ht, pst);

        }

        private static Detail GetVPSampleDetails(IntPtr hw)
        {
            clickVoiceAndZonesTab(hw, 10);
            return new Detail
            {
                mute = GetBooleanCtrlValue(hw, 1),
                solo = GetBooleanCtrlValue(hw, 2),
                groop = GetEditValue(hw, 2).ToInt(),
                sample = GetComboboxValue(hw, 9),
                direction = GetComboboxValue(hw, 10),
                notePlacement = new LowHighPlacement
                {
                    low = GetEditValue(hw, 6),
                    fadel = GetEditValue(hw, 7).ToInt(),
                    high = GetEditValue(hw, 8),
                    fadeh = GetEditValue(hw, 9).ToInt(),
                },

                orK = GetEditValue(hw, 10),
                volume = GetEditValue(hw, 11).ToFloat(),
                pan = GetEditValue(hw, 12).ToInt(),
                xpoz = GetEditValue(hw, 13).ToInt(),
                coarse = GetEditValue(hw, 14).ToInt(),
                fine = GetEditValue(hw, 15).ToFloat(),
            };
        }

        static Zone GetSelectionZone(IntPtr hw)
        {
            return new Zone
            {
                detail = GetVPSampleDetails(hw),
                routing = getVPSampleControlRouting(hw),
            };
        }


        static void ImportLowHighPlacementValues(
            IntPtr hw, 
            ref LowHighPlacement linkPlacement, 
            ref LowHighPlacement currentVPSelPlacement,
            int line
            )
        {
            if ( currentVPSelPlacement.low == "0".ToString()  &&
                !string.IsNullOrEmpty(linkPlacement.low))
            {
                SetEditValue(hw, 263 + line * 4, linkPlacement.low);
            }
            if (currentVPSelPlacement.fadel != null &&
                linkPlacement.fadel != null &&
                (int)currentVPSelPlacement.fadel == 0 
                )
            {
                SetEditValue(hw, 264 + line * 4, linkPlacement.fadel.ToString());
            }
            if (currentVPSelPlacement.high == "127".ToString()  &&
                !string.IsNullOrEmpty(linkPlacement.high) )
            {
                SetEditValue(hw, 265 + line * 4, linkPlacement.high);
            }
            if (currentVPSelPlacement.fadeh != null &&
                linkPlacement.fadeh != null &&
                (int)currentVPSelPlacement.fadeh == 127
                 )
            {
                SetEditValue(hw, 266 + line * 4, linkPlacement.fadeh.ToString());
            }
        }

        static void ImportLowHighPlacementNotes(
            IntPtr hw,
            ref LowHighPlacement linkPlacement,
            ref LowHighPlacement currentVPSelPlacement,
            int line
            )
        {
            if ( currentVPSelPlacement.low == "C-2".ToString() &&
                !string.IsNullOrEmpty(linkPlacement.low))
            {
                SetEditValue(hw, 263 + line * 4, linkPlacement.low);
            }
            if (currentVPSelPlacement.fadel != null &&
                (int)currentVPSelPlacement.fadel == 0 &&
                linkPlacement.fadel != null)
            {
                SetEditValue(hw, 264 + line * 4, linkPlacement.fadel.ToString());
            }
            if ( currentVPSelPlacement.high == "G8".ToString() &&
                !string.IsNullOrEmpty(linkPlacement.high))
            {
                SetEditValue(hw, 265 + line * 4, linkPlacement.high);
            }
            if (currentVPSelPlacement.fadeh != null &&
                (int)currentVPSelPlacement.fadeh == 127 &&
                linkPlacement.fadeh != null)
            {
                SetEditValue(hw, 266 + line * 4, linkPlacement.fadeh.ToString());
            }
        }

        static void ImportLinkControlToVoices(
          IntPtr hw,
          ref SourcePlacement linkPlacement,
          ref SourcePlacement currentVPSelPlacement,
          int ccn
          )
        {
            if(  !string.IsNullOrEmpty(linkPlacement.source))
            {
                SetCBBElement(hw, 539 + ccn - 3, 113 + ccn - 3, linkPlacement.source);
            }
            ImportLowHighPlacementValues(hw, ref linkPlacement.placement, ref currentVPSelPlacement.placement, ccn );

        }


        static void WriteIfDefault(
            IntPtr hw,
            ref ControlRouting linkPlacement,
            ref ControlRouting currentVPSelPlacement
            )
        {

            ImportLowHighPlacementValues(hw, ref linkPlacement.Key, ref currentVPSelPlacement.Key, 0);
            ImportLowHighPlacementValues(hw, ref linkPlacement.Vel, ref currentVPSelPlacement.Vel, 1);
            ImportLowHighPlacementValues(hw, ref linkPlacement.RT, ref currentVPSelPlacement.RT, 2);

            ImportLinkControlToVoices(hw, ref linkPlacement.CC1, ref currentVPSelPlacement.CC1, 3);            
            ImportLinkControlToVoices(hw, ref linkPlacement.CC2, ref currentVPSelPlacement.CC2, 4);
            ImportLinkControlToVoices(hw, ref linkPlacement.CC3, ref currentVPSelPlacement.CC3, 5);
            ImportLinkControlToVoices(hw, ref linkPlacement.CC4, ref currentVPSelPlacement.CC4, 6);
            ImportLinkControlToVoices(hw, ref linkPlacement.CC5, ref currentVPSelPlacement.CC5, 7);
        }

        static void WriteIfDefault(
           IntPtr hw,
           ref LinkData linkPlacement,
           ref Zone currentVPSelPlacement
           )
        {
            if (linkPlacement.volume != null
                && currentVPSelPlacement.detail.volume != null
                )
            {
                SetEditValue(hw, 11, linkPlacement.volume.ToString());
            };

            if (linkPlacement.pan != null
               && currentVPSelPlacement.detail.pan != null
               )
            {
                SetEditValue(hw, 12, linkPlacement.pan.ToString());
            };

            if (linkPlacement.xpoz != null
               && currentVPSelPlacement.detail.xpoz != null
               )
            {
                SetEditValue(hw, 13, linkPlacement.xpoz.ToString());
            };

            if (linkPlacement.fine != null
               && currentVPSelPlacement.detail.fine != null
               )
            {
                SetEditValue(hw, 14, linkPlacement.fine.ToString());
            }
            WriteIfDefault(hw, ref linkPlacement.routing, ref currentVPSelPlacement.routing);
        }

        private static void CopyLinkDataToVP(IntPtr hw, IntPtr ht, ref LinkData l, ref Zone z, int pstnum)
        {
            selectNthPresetVoicsAndZone(hw, ht, pstnum);
            //var d = getVPSampleDetails(hw);
            clickVoiceAndZonesTab(hw,10);
            WriteIfDefault(hw, ref l, ref z);
        }
        static SourcePlacement getDetailCCx(IntPtr hw, int n)
        {
            return new SourcePlacement { 
                source  = GetEditValue(hw, 539 + n),
                placement = getHighLowDetail(hw, n + 3)
            };
        }

        static LowHighPlacement getHighLowDetail(IntPtr hw , int line)
        {
            return new LowHighPlacement
            {
                low = GetEditValue(hw, 263 + line * 4),
                fadel = GetEditValue(hw, 264 + line * 4).ToInt(),
                high = GetEditValue(hw, 265 + line * 4),
                fadeh = GetEditValue(hw, 266 + line * 4).ToInt(),
            };
        }
        
        private static ControlRouting getVPSampleControlRouting(IntPtr hw)
        {
            return new ControlRouting
            {
                Key = getHighLowDetail(hw,0),
                Vel = getHighLowDetail(hw, 1),
                RT  = getHighLowDetail(hw, 2),
                CC1 = getDetailCCx(hw, 0),
                CC2 = getDetailCCx(hw, 1),
                CC3 = getDetailCCx(hw, 2),
                CC4 = getDetailCCx(hw, 3),
                CC5 = getDetailCCx(hw, 4),
            };
        }


        private static Zone getVPSampleZone(IntPtr hw)
        {
            return new Zone
            {
                detail = GetVPSampleDetails(hw),
                routing = getVPSampleControlRouting(hw),    
            };
        }

        private static void DeletePreset(IntPtr hw, IntPtr ht, int from, int num = 0)
        {
            selectPresetView(hw, ht);
            var hl = AutoItX.ControlGetHandle(hw, "[CLASS:SysListView32;INSTANCE:4]");
            var hh = AutoItX.ControlGetHandle(hw, "[CLASS:SysHeader32;INSTANCE:4]");
            if (!((int)hl > 0 && (int)hh > 0))
            {
                throw new Exception("no listview!!!");
            }
            var s = AutoItX.ControlListView(
                       hw,
                       hl,
                        "SelectClear",
                        "",
                        ""
                        );
            if (s != "1")
            {
                throw new Exception("command SelectClear failed!!!");
            }
            s = AutoItX.ControlListView(
                          hw,
                          hl,
                           "Select",
                           $"{from}",
                           num == 0 ? "" : $"{from + num}"
                           );
            if (s != "1")
            {
                throw new Exception("command Select failed!!!");
            }
            AutoItX.ControlClick(hw, hh, "right");
            AutoItX.ControlSend(hw, hh, "{DOWN 5}{ENTER}");
        }


        static string ToFring(this float str)
        {
            return str.ToString(CultureInfo.InvariantCulture);
        }

        static int? ToPresetNumber(this string str)
        {
            if (str != String.Empty)
            {
                str = str.Remove(0, 1);
                str = str.Remove(4, str.Length - 4);
                return str.ToInt();
            }
            return null;
        }

        static int? ToInt(this string str)
        {
            if (String.IsNullOrEmpty(str)) return null;
            return int.Parse(str,
                NumberStyles.Integer 
                | NumberStyles.AllowTrailingSign 
                | NumberStyles.AllowTrailingWhite);
        }

        static float ToFloat(this string str)
        {
            if (String.IsNullOrEmpty(str)) return 0f;
            return float.Parse(str,
                NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture);
        }

        static string selectNthFromM(IntPtr hw, IntPtr ht, int m, int n)
        {
            return AutoItX.ControlTreeView(
                   hw,
                   ht,
                    "Select",
                    $"#{m}|#{n}",
                    ""
                );
        }

        static string select_MNO(IntPtr hw, IntPtr ht, int m, int n, int o)
        {
            var s = AutoItX.ControlTreeView(
                   hw,
                   ht,
                    "Select",
                    $"#{m}|#{n}|#{o}",
                    ""
                );
            return s;
        }
        static int GetSelectedNode(IntPtr hw, IntPtr ht)
        {
            var sl = "#0|#".Length;
            var resp = AutoItX.ControlTreeView(  hw,  ht,  "GetSelected", "1", "" ); //"1" is for useindex
            if (resp.Length < sl) return -1;
            resp = resp.Remove(0, sl);
            int ret = -1;
            int.TryParse(resp, out ret);
            return ret;
        }



        //N time the first letter
        static void SetComboboxElement(IntPtr hw, int ctrln, char el, int repeat)
        {
            var str = ClassName_cbb_text(ctrln);
            var chdl = AutoItX.ControlGetHandle(hw, str);
            var rzClick = AutoItX.ControlClick(hw, chdl, "left", 1, 0, 0);
            //N times the letter to access the correct element ... :/
            var cmd = $"{{{el} {repeat}}}" + "{ENTER}";
            AutoItX.ControlSend(
                   hw,
                   chdl,
                   cmd
                );
        }
        //Element overall Position
        static void SetComboboxElement(IntPtr hw, int ctrlN, int elN)
        {
            var str = ClassName_cbb_slct(ctrlN);
            var hc = AutoItX.ControlGetHandle(hw, str);
            var rzClick = AutoItX.ControlClick(hw, hc);
            var rnc2 = AutoItX.ControlSend(hw, hc, $"{{DOWN {(elN + 1).ToString()}}}");
            var rnc = AutoItX.ControlSend(hw, hc, $"{{ENTER}}");
        }

        static void selectGroup(IntPtr hw, IntPtr ht, int n, int gn)
        {
            selectNthPresetVoiceProc(hw, ht, n);

            selectNthFromM(hw, ht, 0, n);
        }
        static string selectPresetView(IntPtr hw, IntPtr ht)
        {
            return AutoItX.ControlTreeView(
                               hw,
                               ht,
                                "Select",
                                $"#0",
                                ""
                            );
        }

        static string selectNthPreset(IntPtr hw, IntPtr ht, int n)
        {
            return selectNthFromM(hw, ht, 0, n);
        }
        static string selectNthSample(IntPtr hw, IntPtr ht, int n)
        {
            return selectNthFromM(hw, ht, 1, n);
        }

        static string selectNthMultisetup(IntPtr hw, IntPtr ht, int n)
        {
            return selectNthFromM(hw, ht, 2, n);
        }
        static string selectNthPresetVoicsAndZone(IntPtr hw, IntPtr ht, int n)
        {
            return select_MNO(hw, ht, 0, n, 0);
        }
        public static string selectNthPresetVoiceLinks(IntPtr hw, IntPtr ht, int n)
        {
            return select_MNO(hw, ht, 0, n, 1);
        }
        public static void selectNthPresetVoiceProc(IntPtr hw, IntPtr ht, int n)
        {
            select_MNO(hw, ht, 0, n, 2);
        }


        static int PasteTime()
        {
            int dupPos = -1;
            var popwh = AutoItX.WinGetHandle("Paste");
            if ((int)popwh > 0)
            {

                var rbh = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:3]");
                var selh = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:4]");
                if (!((int)rbh > 0 && (int)selh > 0))
                {
                    throw new Exception("PasteDialog: buttons no found");
                }
                var posCtrlH = AutoItX.ControlGetHandle(popwh, GetEditClassName(1));
                if (!((int)posCtrlH > 0))
                {
                    throw new Exception("PasteDialog: target control not found");
                }
                dupPos = int.Parse(GetEditValue(popwh, 1));

                AutoItX.ControlClick(popwh, rbh);   //dest will be at suggested prese 
                AutoItX.ControlClick(popwh, selh); //deselect move preset option
                var okbutt = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:7]");
                if (!((int)okbutt > 0))
                {
                    throw new Exception("PasteDialog: buttons no found");
                }
                AutoItX.ControlClick(popwh, okbutt);
            }
            return dupPos;
        }
        public static int DuplicatePresets(IntPtr hw, IntPtr ht, int from, int num = 0)
        {
            int dupPos = -1;
            selectPresetView(hw, ht);
            var hl = AutoItX.ControlGetHandle(hw, "[CLASS:SysListView32;INSTANCE:4]");
            var hh = AutoItX.ControlGetHandle(hw, "[CLASS:SysHeader32;INSTANCE:4]");
            if (!((int)hl > 0 && (int)hh > 0))
            {
                throw new Exception("no listview!!!");
            }
            var s = AutoItX.ControlListView(
                       hw,
                       hl,
                        "SelectClear",
                        "",
                        ""
                        );
            if (s != "1")
            {
                throw new Exception("command SelectClear failed!!!");
            }
            s = AutoItX.ControlListView(
                          hw,
                          hl,
                           "Select",
                           $"{from}",
                           num == 0 ? "" : $"{from + num}"
                           );
            if (s != "1")
            {
                throw new Exception("command Select failed!!!");
            }
            var t = false;
            while (!t)
            {
                AutoItX.ControlClick(hw, hh, "right");
                AutoItX.ControlSend(hw, hh, "{DOWN 3}{ENTER}");
                AutoItX.ControlClick(hw, hh, "right");
                AutoItX.ControlSend(hw, hh, "c{ENTER}");
                AutoItX.ControlClick(hw, hh, "right");
                AutoItX.ControlSend(hw, hh, "{DOWN 4}{ENTER}");

                dupPos = PasteTime();
                t = dupPos > 0;
            }
            return dupPos;

        }

        static string getEditElement(int el)
        {
            return "Edit" + el.ToString();
        }


        static void setFilterToDesigner(IntPtr hw, IntPtr ht, int hex1, int hex2, string t)
        {
        }
        static string GetEditClassName(int v)
        {
            return $"[CLASS:Edit; INSTANCE: {v}]";
        }


        static string ClassName_Button(int v)
        {
            return $"[CLASS:Button; INSTANCE:{v}]";
        }

        static string ClassName_cbb_text(int v)
        {
            return $"[CLASS:Afx:0000000140000000:8b:0000000000010003:0000000000000000:0000000000000000; INSTANCE:{v}]";
        }
        static string ClassName_cbb_slct(int v)
        {
            return $"[CLASS:Afx:0000000140000000:83:0000000000010003:0000000000000000:0000000000000000; INSTANCE:{v}]";
        }

        static string ClassName_glob()
        {
            return $"[CLASS:Afx:0000000140000000:8:0000000000010003:0000000000000006:000000000001002B; INSTANCE:1]";
        }

        public static void ApplySource2DstRange(int src, uint dstart, uint dend)
        {
            if (dstart < dend) return;
            for (uint i = dstart; i < dend; ++i)
            {

            }
        }


        static void handlingWindow(string wn)
        {
            AutoItX.AutoItSetOption("WinTitleMatchMode", 2);
            var hw = AutoItX.WinGetHandle(wn);
            if (hw != null && (int)hw > 0)
                Console.WriteLine(hw.ToString());

            //var p =  AutoItX.WinGetProcess("EmulatorX.exe");
            var ht = AutoItX.ControlGetHandle(hw, "[CLASS:SysTreeView32;INSTANCE:1]");
            if ((int)ht > 0)
            {
                Console.WriteLine(ht.ToString());
                int i = 0;
                while (true)
                {
                    selectNthPresetVoiceLinks(hw, ht, i++);
                    //selectNthSample(hw, ht, 0);
                    //Thread.Sleep(1000);
                    //selectNthMultisetup(hw, ht, 0);
                    //Thread.Sleep(1000);
                }
            }
        }
        private static void SelectNextGroup(IntPtr hw, IntPtr ht, int pnum, string txt)
        {
            selectNthPresetVoiceProc(hw, ht, pnum);
            var str1 = ClassName_cbb_slct(119);
            var h1 = AutoItX.ControlGetHandle(hw, str1);
            var r1 = AutoItX.ControlSend(hw, h1, "{DOWN} {ENTER}");
            return;
        }
        static string setToNextValue(IntPtr hw, IntPtr hc, IntPtr ht, char c)
        {
            var orig = AutoItX.ControlGetText(hw, ht);

            //AutoItX.ControlClick(hw, hc);
            while (true)
                AutoItX.ControlSend(hw, hc, $"{c}");
            return AutoItX.ControlGetText(hw, hc);
        }

        static string selectNextDownValue(IntPtr hw, IntPtr hc, IntPtr ht)
        {

            AutoItX.ControlSend(hw, hc, "{DOWN}");
            var t = AutoItX.ControlGetText(hw, ht);
            return t;
        }


        static string selectNextUpValue(IntPtr hw, IntPtr hc, IntPtr ht)
        {

            AutoItX.ControlSend(hw, hc, "{UP}");
            var t = AutoItX.ControlGetText(hw, ht);
            return t;
        }

        static string selectFirstValue(IntPtr hw, IntPtr hc, IntPtr ht)
        {
            var rzClick = AutoItX.ControlClick(hw, hc, "left", 1, 0, 0);
            AutoItX.ControlSend(hw, hc, "{DOWN}{ENTER}");
            var t = AutoItX.ControlGetText(hw, ht);
            return t;
        }
        static string selectLastValue(IntPtr hw, IntPtr hc, IntPtr ht)
        {

            var rzClick = AutoItX.ControlClick(hw, hc, "left", 1, 0, 0);
            AutoItX.ControlSend(hw, hc, "{UP}{ENTER}");
            var t = AutoItX.ControlGetText(hw, ht);
            return t;
        }

        private static void SelectElement(IntPtr hw, IntPtr ht, IntPtr hc, string txt)
        {
            var t = AutoItX.ControlGetText(hw, ht);
            if (t == txt) return;
            var lv = selectLastValue(hw, hc, ht);
            if (lv == txt) return;
            var fv = selectFirstValue(hw, hc, ht);
            if (fv == txt) return;
            do
            {
                t = selectNextDownValue(hw, hc, ht);
                if (t == txt) return;
            }
            while (t != txt && t != lv);
            //recommencer si pour x raisons l'element n'a pas été trouvé
            if (t == lv)
                SelectElement(hw, ht, hc, txt);
        }


        public static void SetCBBElement(IntPtr hw, int cbbtxt, int cbbctrl, string txt)
        {
            var str2 = ClassName_cbb_text(cbbtxt);
            var h2 = AutoItX.ControlGetHandle(hw, str2);


            var str1 = ClassName_cbb_slct(cbbctrl);
            var h1 = AutoItX.ControlGetHandle(hw, str1);


            SelectElement(hw, h2, h1, txt);

        }
        private static void SelectGroup(IntPtr hw, IntPtr ht, int pnum, string txt)
        {
            selectNthPresetVoiceProc(hw, ht, pnum);
            SetCBBElement(hw, 244, 119, txt);

        }

        public static bool presetExists(IntPtr hw, IntPtr ht, int n)
        {
            var str = AutoItX.ControlTreeView(
                   hw,
                   ht,
                    "Exists",
                    $"#0#" + n.ToString(),
                    ""
                );
            Console.WriteLine(str);
            str = AutoItX.ControlTreeView(
                  hw,
                  ht,
                   "Checked",
                   $"#0#" + n.ToString(),
                   ""
               );
            str = AutoItX.ControlTreeView(
                  hw,
                  ht,
                   "Select",
                   $"#0#" + n.ToString(),
                   ""
               );
            str = AutoItX.ControlTreeView(
                  hw,
                  ht,
                   "GetSelected",
                   "1",
                   ""
               );
            Console.WriteLine(str);
            return str == "1" ? true : false;
        }



        public static int getMaxPresetNumber(IntPtr hw, IntPtr ht)
        {
            var str = AutoItX.ControlTreeView(
                   hw,
                   ht,
                    "GetItemCount",
                    $"#0",
                    ""
                );
            var ret = int.Parse(str);
            return ret;
        }

        static void SelectGroup(IntPtr hw, IntPtr ht, int pnum, int gnum)
        {
            SelectGroup(hw, ht, pnum, gnum.ToString());
        }
        public static int GetSelectedInstrumentNum(IntPtr hw, IntPtr ht)
        {
            return GetSelectedNode(hw, ht);
        }
      

        public static void UnminimizeWindow(IntPtr hw)
        {
            AutoItX.WinSetOnTop(hw, 1);

            WindowsUIF.RECT rc;
            WindowsUIF.GetWindowRect(hw, out rc);
            //Screenshooter.MoveWindow(hw, 0, 0, rc.Width, rc.Height, false);
            AutoItX.WinSetState(hw,
                AutoItX.SW_SHOWMAXIMIZED
                );

        }

        public static void SelectAllVoices(IntPtr hw, IntPtr ht, int pnum)
        {
            SelectGroup(hw, ht, pnum, "All");
        }



        public static Tuple<IntPtr, IntPtr> getHandles()
        {

            AutoItX.AutoItSetOption("WinTitleMatchMode", 2); //substring
            AutoItX.AutoItSetOption("PixelCoordMode", 1); //substring
            IntPtr hw = AutoItX.WinGetHandle(" - Emulator X");
            if ((int)hw > 0)
            {
                //Console.WriteLine(hw);
                var nht = WindowHandleInfo.GetEXHandle_treeview(hw, 1);
                var ht = AutoItX.ControlGetHandle(hw, "[CLASS:SysTreeView32;INSTANCE:1]");

                return new Tuple<IntPtr, IntPtr>(hw, ht);
            }
            return new Tuple<IntPtr, IntPtr>(new IntPtr(0), new IntPtr(0));

        }


        static IntPtr getControlHandle(IntPtr hw, int num)
        {
            var str = ClassName_cbb_text(num);
            return AutoItX.ControlGetHandle(hw, str);
        }
        static void clickVPTab(IntPtr hw, int num)
        {
            var h = getControlHandle(hw, 48);

            var pos = AutoItX.ControlGetPos(hw, h);
            int x = num * (pos.Width / 7);
            //int y = -pos.Height / 2;
            var rzClick = AutoItX.ControlClick(hw, h, "left", 1, x, 0);
        }

     
        static string getControlValue(IntPtr hw, int num)
        {
            var h = getControlHandle(hw, num);
            return AutoItX.ControlGetText(hw, h);
        }
        public static void PutControlInVisibleZone(IntPtr hw, IntPtr hCtrl)
        {
            var wp = AutoItX.WinGetPos(hw);
            var lv = AutoItX.ControlGetPos(hw, hCtrl);

        }
        public static bool GetBooleanCtrlValue(IntPtr hw, IntPtr hCtrl, int pnum)
        {
            var lv = AutoItX.ControlGetPos(hw, hCtrl);
            AutoItX.ControlEnable(hw, hCtrl); //in case it's greyed
            var x = (int)(lv.Width * .5);
            var y = (int)(lv.Height * .5);
            var col = WindowsUIF.GetPixelColor(hCtrl, x, y).ToArgb();
            //Console.WriteLine(col);
            return col == -13904170;
        }
        public static bool GetBooleanCtrlValue(IntPtr hw, int pnum)
        {
            var hCtrl = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(pnum));
            return GetBooleanCtrlValue(hw, hCtrl, pnum);
        }

        public static void SetBooleanCtrlValue(IntPtr hw, int pnum, bool data)
        {
            var hCtrl = AutoItX.ControlGetHandle(hw, ClassName_cbb_slct(pnum));
            if (GetBooleanCtrlValue(hw, hCtrl, pnum) != data)
            {
                AutoItX.ControlClick(hw, hCtrl);
            }
        }

        private static int getFGEndStep(IntPtr hw, int n)
        {
            var sret = GetEditValue(hw, 77 + n * 4);
            int ret = -1;

            int.TryParse(sret, out ret);
            return ret;
        }
 

        private static string GetComboboxValue(IntPtr hw, int pnum)
        {

            var str = ClassName_cbb_text(pnum);
            var bar = AutoItX.ControlGetHandle(hw, str);
            return AutoItX.ControlGetText(hw, bar);
        }
        private static int SetComboboxValue(IntPtr hw, int pnum, string txt)
        {
            var str = ClassName_cbb_slct(pnum);
            var ctrlH = AutoItX.ControlGetHandle(hw, str);
            var rzClick = AutoItX.ControlClick(hw, ctrlH, "left", 1, 0, 0);
            return AutoItX.ControlSend(hw, ctrlH, txt);
        }


        public static string GetEditValue(IntPtr hw, int num)
        {
            var str = GetEditClassName(num);
            var chdl = AutoItX.ControlGetHandle(hw, str);
            return AutoItX.ControlGetText(hw, chdl);

        }
        public static int SetEditValue(IntPtr hw, int num, string data)
        {
            var str = GetEditClassName(num);
            var chdl = AutoItX.ControlGetHandle(hw, str);
            AutoItX.ControlSetText(hw, chdl, data);
            return AutoItX.ControlSend(hw, chdl, "{ENTER}");
        }
        //public static int SetEditIntValue(IntPtr hw, int num, int data)
        //{
        //    var str = GetEditClassName(num);
        //    var chdl = AutoItX.ControlGetHandle(hw, str);
        //    var ret = AutoItX.ControlSend(hw, chdl, data.ToString());
        //    return ret;

        //}

        #region
        //START VP GETTERS
        public static Tune getVP_tune(IntPtr hw)
        {
            return new Tune
            {
                transpose = GetEditValue(hw, 2),
                Fine = GetEditValue(hw, 3),
                Coarse = GetEditValue(hw, 4),
            };
        }
        public static Chorus getVP_Chorus(IntPtr hw)
        {
            return new Chorus
            {
                amount = GetEditValue(hw, 12),
                width = GetEditValue(hw, 13),
                initms = GetEditValue(hw, 14),
            };
        }
        public static Twistaloop getVP_Twistaloop(IntPtr hw)
        {
            return new Twistaloop
            {
                SpeedPrcnt = GetEditValue(hw, 5),
                SoopPrcnt = GetEditValue(hw, 6),
                StartAtLoop = GetBooleanCtrlValue(hw, 2)
            };
        }
        public static Bend GetVP_Bend(IntPtr hw)
        {
            return new Bend
            {
                Up = GetEditValue(hw, 7),
                Dn = GetEditValue(hw, 8),
            };
        }


        public static VPKey GetVP_Key(IntPtr hw)
        {
            return new VPKey
            {
                assignGroup = GetComboboxValue(hw, 16),
                keymode = GetComboboxValue(hw, 17),
                bpm = GetBooleanCtrlValue(hw, 3),
                delay = GetEditValue(hw, 9),
                latch = GetBooleanCtrlValue(hw, 4),
                offset = GetEditValue(hw, 10).ToFloat(),
            };
        }
        public static Glide GetVP_Glide(IntPtr hw)
        {
            return new Glide
            {
                Rate = GetEditValue(hw, 11).ToFloat(),
                Curve = GetComboboxValue(hw, 18),
            };
        }
        public static Env GetVP_Env(IntPtr hw, int n)
        {
            return new Env
            {
                bpm = GetBooleanCtrlValue(hw, 17 + n * 2),
                Attack1Time = GetEditValue(hw, 40 + n * 12),
                Attack1Level = GetEditValue(hw, 46 + n * 12),
                Attack2Time = GetEditValue(hw, 41 + n * 12),
                Attack2Level = GetEditValue(hw, 47 + n * 12),
                Decay1Time = GetEditValue(hw, 42 + n * 12),
                Decay1Level = GetEditValue(hw, 48 + n * 12),
                Decay2Time = GetEditValue(hw, 43 + n * 12),
                Decay2Level = GetEditValue(hw, 49 + n * 12),
                Release1Time = GetEditValue(hw, 44 + n * 12),
                Release1Level = GetEditValue(hw, 50 + n * 12),
                Release2Time = GetEditValue(hw, 45 + n * 12),
                Release2Level = GetEditValue(hw, 51 + n * 12),
            };
        }
        public static Env GetVP_AmpEnv(IntPtr hw)
        {
            return GetVP_Env(hw, 0);
            return new Env
            {
                bpm = GetBooleanCtrlValue(hw, 17),
                Attack1Time = GetEditValue(hw, 40),
                Attack1Level = GetEditValue(hw, 46),
                Attack2Time = GetEditValue(hw, 41),
                Attack2Level = GetEditValue(hw, 47),
                Decay1Time = GetEditValue(hw, 42),
                Decay1Level = GetEditValue(hw, 48),
                Decay2Time = GetEditValue(hw, 43),
                Decay2Level = GetEditValue(hw, 49),
                Release1Time = GetEditValue(hw, 44),
                Release1Level = GetEditValue(hw, 50),
                Release2Time = GetEditValue(hw, 45),
                Release2Level = GetEditValue(hw, 51),
            };
        }
        public static Env GetVP_FilterEnv(IntPtr hw)
        {
            return GetVP_Env(hw, 1);

            return new Env
            {
                bpm = GetBooleanCtrlValue(hw, 19),
                Attack1Time = GetEditValue(hw, 52),
                Attack1Level = GetEditValue(hw, 58),
                Attack2Time = GetEditValue(hw, 53),
                Attack2Level = GetEditValue(hw, 59),
                Decay1Time = GetEditValue(hw, 54),
                Decay1Level = GetEditValue(hw, 60),
                Decay2Time = GetEditValue(hw, 55),
                Decay2Level = GetEditValue(hw, 61),
                Release1Time = GetEditValue(hw, 56),
                Release1Level = GetEditValue(hw, 62),
                Release2Time = GetEditValue(hw, 57),
                Release2Level = GetEditValue(hw, 63),
            };
        }
        public static Env GetVP_AuxEnv(IntPtr hw)
        {
            return GetVP_Env(hw, 2);
            return new Env
            {
                bpm = GetBooleanCtrlValue(hw, 21),
                Attack1Time = GetEditValue(hw, 64),
                Attack1Level = GetEditValue(hw, 70),
                Attack2Time = GetEditValue(hw, 71),
                Attack2Level = GetEditValue(hw, 59),
                Decay1Time = GetEditValue(hw, 72),
                Decay1Level = GetEditValue(hw, 60),
                Decay2Time = GetEditValue(hw, 73),
                Decay2Level = GetEditValue(hw, 61),
                Release1Time = GetEditValue(hw, 74),
                Release1Level = GetEditValue(hw, 62),
                Release2Time = GetEditValue(hw, 75),
                Release2Level = GetEditValue(hw, 63),
            };
        }

        public static LFO GetVP_LFO1(IntPtr hw)
        {
            return new LFO
            {
                keySync = GetBooleanCtrlValue(hw, 12),
                bpm = GetBooleanCtrlValue(hw, 11),
                Frequency = GetEditValue(hw, 32),
                Delay = GetEditValue(hw, 33),
                Variation = GetEditValue(hw, 34),
            };
        }

        public static LFO GetVP_LFO2(IntPtr hw)
        {
            return new LFO
            {
                keySync = GetBooleanCtrlValue(hw, 15),
                bpm = GetBooleanCtrlValue(hw, 14),
                Frequency = GetEditValue(hw, 35),
                Delay = GetEditValue(hw, 36),
                Variation = GetEditValue(hw, 37),
            };
        }

        static Step[] GetSteps(IntPtr hw, int fgnum)
        {
            // cannot direct query: select the stepnum first
            var trets = new List<Step>();
            var bar = AutoItX.ControlGetHandle(hw, getEditElement(78 + 4 * fgnum));
            //var bar = AutoItX.ControlGetHandle(hw, ClassName_cbb_text(114 + 4 * fgnum));
            for (int i = 0; i < 64; ++i)
            {
                //AutoItX.ControlFocus(hw, bar);
                AutoItX.ControlSetText(hw, bar, (i + 1).ToString());
                AutoItX.ControlSend(hw, bar, "{ENTER 1}");
                trets.Add(new Step
                {
                    value = GetEditValue(hw, 79 + 4 * fgnum).ToFloat(),
                    trigger = GetBooleanCtrlValue(hw, 27 + 8 * fgnum),
                });
            }
            //Console.WriteLine(trets[0]);
            //Console.WriteLine(trets[63]);
            return trets.ToArray();
        }



        public static FuncGen GetFuncGen(IntPtr hw, int n)
        {
            clickVPTab(hw, 4 + n);
            return new FuncGen
            {
                StepRate = GetEditValue(hw, 76 + n * 4),
                bpm = GetBooleanCtrlValue(hw, 23 + n * 8),
                sync = GetComboboxValue(hw, 111 + n * 8),
                smooth = GetBooleanCtrlValue(hw, 25 + n * 8),
                direction = GetComboboxValue(hw, 112 + n * 8),
                endStep = int.Parse(GetEditValue(hw, 77 + n * 4)),
                steps = GetSteps(hw, n)
            };
        }


        static Cord GetCord(IntPtr hw, int n)
        {
            return new Cord()
            {
                source = GetComboboxValue(hw, 136 + (n * 3)),
                destination = GetComboboxValue(hw, 138 + (n * 3)),
                value = GetEditValue(hw, 88 + n).ToFloat()
            };
        }
        public static Cord[] GetVP_Cords(IntPtr hw)
        {
            var tret = new List<Cord>();
            for (int i = 0; i < 36; ++i)
            {
                tret.Add(GetCord(hw, i));
            }
            return tret.ToArray();
        }

        static FilterMorpher GetMorpher(IntPtr hw)
        {
            return new FilterMorpher()
            {
                left = new MorphParameters()
                {
                    lofreq = GetEditValue(hw, 17).ToFloat(),
                    hifreq = GetEditValue(hw, 18).ToFloat(),
                    gain = GetEditValue(hw, 19).ToFloat()
                },
                right = new MorphParameters()
                {
                    lofreq = GetEditValue(hw, 20).ToFloat(),
                    hifreq = GetEditValue(hw, 21).ToFloat(),
                    gain = GetEditValue(hw, 22).ToFloat()
                },
            };
        }

        static void SetMorpher(IntPtr hw, FilterMorpher m)
        {
            SetEditValue(hw, 17, (m.left.lofreq).ToFring());
            SetEditValue(hw, 18, (m.left.hifreq).ToFring());
            SetEditValue(hw, 19, (m.left.gain).ToFring());

            SetEditValue(hw, 20, (m.right.lofreq).ToFring());
            SetEditValue(hw, 21, (m.right.hifreq).ToFring());
            SetEditValue(hw, 22, (m.right.gain).ToFring());
        }
        static FilterDesigner GetDesigner(IntPtr hw)
        {
            var ret = new FilterDesigner();
            var lst = new List<FilterDesignerStage>();
            for (int i = 0; i < 6; ++i)
            {
                SetEditValue(hw, 23, (i + 1).ToString());
                var typ = GetComboboxValue(hw, 38);
                var f1 = 0.0f;
                var f2 = 0.0f;
                var g1 = 0.0f;
                var g2 = 0.0f;

                if (typ != "Off")
                {
                    f1 = GetEditValue(hw, 18).ToFloat();
                    f2 = GetEditValue(hw, 19).ToFloat();
                    g1 = GetEditValue(hw, 21).ToFloat();
                    g2 = GetEditValue(hw, 22).ToFloat();
                }
                lst.Add(new FilterDesignerStage()
                {
                    freq1 = f1,
                    freq2 = f2,
                    ftype = typ,
                    gain1 = g1,
                    gain2 = g2,
                });
            };
            ret.stages = lst.ToArray();
            return ret;
        }
        static void SetFilterStage(IntPtr hw, int stageNum, FilterDesignerStage stage)
        {
            SetEditValue(hw, 23, (stageNum + 1).ToString());

            SetCBBElement(hw, 38, 9, stage.ftype);
            SetEditValue(hw, 18, (stage.freq1).ToFring());
            SetEditValue(hw, 19, (stage.gain1).ToFring());
            SetEditValue(hw, 21, (stage.freq2).ToFring());
            SetEditValue(hw, 22, (stage.gain2).ToFring());
        }
        static void SetDesigner(IntPtr hw, FilterDesigner dez)
        {
            for (int i = 0; i < 6; ++i)
            {
                SetFilterStage(hw, i, dez.stages[i]);
            }
        }


        public static Filter getFilter(IntPtr hw)
        {
            var ret = new Filter();


            ret.name = GetComboboxValue(hw, 28);
            ret.freq = GetEditValue(hw, 15).ToFloat();
            ret.rez = GetEditValue(hw, 16).ToFloat();

            if (ret.name == "Morph Designer")
            {
                ret.designer = GetDesigner(hw);
                ret.morpher = null;
            }
            else if (ret.name.Contains("Morph"))
            {
                ret.morpher = GetMorpher(hw);
                ret.designer = null;
            }
            else
            {
                ret.morpher = null;
                ret.designer = null;
            }

            return ret;

        }

        public static Amplifier GetAmplifier(IntPtr hw)
        {
            return new Amplifier()
            {
                Volume = GetEditValue(hw, 24).ToFloat(),
                Pan = GetEditValue(hw, 25).ToFloat(),
                FxWetDry = GetEditValue(hw, 26).ToFloat(),
                AmpEnvRange = GetEditValue(hw, 27).ToFloat(),
                Main = int.Parse(GetEditValue(hw, 28)),
                Aux1 = int.Parse(GetEditValue(hw, 29)),
                Aux2 = int.Parse(GetEditValue(hw, 30)),
                Aux3 = int.Parse(GetEditValue(hw, 31)),
                ClassicResponse = GetBooleanCtrlValue(hw, 10),
            };
        }


        static void clickVoiceAndZonesTab(IntPtr hw, int num)
        {
            
            var h = getControlHandle(hw, 1);
            var pos = AutoItX.ControlGetPos(hw, h);
            int x = num * (pos.Width / 11);
            var rzClick = AutoItX.ControlClick(hw, h, "left", 1, x, 0);
        }
        static void clickCCWinTab(IntPtr hw, int num)
        {
            var h = getControlHandle(hw, 1);
            var pos = AutoItX.ControlGetPos(hw, h);
            int x = num * (pos.Width / 8);
            var rzClick = AutoItX.ControlClick(hw, h, "left", 1, x, 0);
        }
        static SourcePlacement getCCWin(IntPtr hw, int n)
        {
            clickCCWinTab(hw, 3 + n);
           
            return new SourcePlacement
            {
                source = GetComboboxValue(hw, 21),
                placement = new LowHighPlacement
                {
                    low = GetEditValue(hw, 14),
                    fadel = int.Parse(GetEditValue(hw, 15)),
                    high = GetEditValue(hw, 16),
                    fadeh = int.Parse(GetEditValue(hw, 17)),
                }
            };
        }

        public static LinkData Get1stLinkData(IntPtr hw )
        {
            var pn = GetComboboxValue(hw, 5).ToPresetNumber();
            if (pn == null) return new LinkData();
            return new LinkData()
            {
                presetNumber  = pn,

                volume = GetEditValue(hw, 2).ToFloat(),
                pan = GetEditValue(hw, 3).ToInt(),
                xpoz = GetEditValue(hw, 4).ToInt(),
                fine = GetEditValue(hw, 5).ToFloat(),

                routing = new ControlRouting
                {

                    Key = new LowHighPlacement
                    {
                        low = GetEditValue(hw, 6),
                        fadel = GetEditValue(hw, 7).ToInt(),
                        high = GetEditValue(hw, 8),
                        fadeh = GetEditValue(hw, 9).ToInt(),
                    },
                    Vel = new LowHighPlacement
                    {
                        low = GetEditValue(hw, 10),
                        fadel = int.Parse(GetEditValue(hw, 11)),
                        high = GetEditValue(hw, 12),
                        fadeh = int.Parse(GetEditValue(hw, 13)),
                    },
                    RT = new LowHighPlacement
                    {
                        low = GetEditValue(hw, 10),
                        fadel = int.Parse(GetEditValue(hw, 11)),
                        high = GetEditValue(hw, 12),
                        fadeh = int.Parse(GetEditValue(hw, 13)),
                    },
                    CC1 = getCCWin(hw, 0),
                    CC2 = getCCWin(hw, 1),
                    CC3 = getCCWin(hw, 2),
                    CC4 = getCCWin(hw, 3),
                    CC5 = getCCWin(hw, 4),
                }
            };
           
        }

        public static void GetLinksSelection_KeyWin(
            IntPtr hw,
            int linknum,
            string low,
            int fadel,
            string high,
            int fadeh
            )
        {
         
        }
        public static void GetLinksSelection_VelWin(
           IntPtr hw,
           int linknum,
           string low,
           int fadel,
           string high,
           int fadeh
           )
        {
            SetEditValue(hw, 10 + 16 * linknum, low);
            SetEditValue(hw, 11 + 16 * linknum, fadel.ToString());
            SetEditValue(hw, 12 + 16 * linknum, high);
            SetEditValue(hw, 13 + 16 * linknum, fadeh.ToString());
        }

        public static void GetSelectionToSplit(
           IntPtr hw,
           int linknum,
           int ccnum,
           string source,
           int low,
           int lowFade,
           int high,
           int highFade
           )
        {

            SetCBBElement(hw, 21 + linknum * 24, 2 + linknum * 2, source);

            SetEditValue(hw, 14 + 16 * linknum, low.ToString());
            SetEditValue(hw, 15 + 16 * linknum, lowFade.ToString());
            SetEditValue(hw, 16 + 16 * linknum, high.ToString());
            SetEditValue(hw, 17 + 16 * linknum, highFade.ToString());

        }


        [DllImport("user32.dll")]
        static extern bool EnableWindow(IntPtr hw, bool enable);


        const uint TPM_LEFTBUTTON = 0x0000;
        const uint TPM_RETURNCMD = 0x0100;
        const uint WM_SYSCOMMAND = 0x0112;

        [DllImport("user32.dll")]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        static extern uint TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public static void ShowContextMenu(IntPtr appWindow, IntPtr myWindow, Point point)
        {
            IntPtr wMenu = GetSystemMenu(appWindow, false);
            AutoItX.ControlClick(appWindow, wMenu);
            // Display the menu
            uint command = TrackPopupMenuEx(wMenu,
                TPM_LEFTBUTTON | TPM_RETURNCMD, (int)point.X, (int)point.Y, myWindow, IntPtr.Zero);
            if (command == 0)
                return;

            PostMessage(appWindow, WM_SYSCOMMAND, new IntPtr(command), IntPtr.Zero);
        }





        public static void ProtectWindow(IntPtr hw)
        {
            EnableWindow(hw, false);
            Macros.UnminimizeWindow(hw);

        }

        public static void UnprotectWindow(IntPtr hw)
        {
            EnableWindow(hw, true);
        }

        public static void SetFilter(IntPtr hw, Filter fltr)
        {
            SetCBBElement(hw, 28, 8, fltr.name);
            SetEditValue(hw, 15, fltr.freq.ToFring());
            SetEditValue(hw, 16, fltr.rez.ToFring());
            if (fltr.name == "Morph Designer")
            {
                SetDesigner(hw, fltr.designer);
            }
            else if (fltr.name.Contains("Morph"))
            {
                SetMorpher(hw, fltr.morpher);
            }


        }

        public static void SetAmplifier(IntPtr hw, Amplifier amp)
        {
            SetEditValue(hw, 24, amp.Volume.ToFring());
            SetEditValue(hw, 25, amp.Pan.ToFring());
            SetEditValue(hw, 26, amp.FxWetDry.ToFring());
            SetEditValue(hw, 27, amp.AmpEnvRange.ToFring());
            SetEditValue(hw, 28, amp.Main.ToString());
            SetEditValue(hw, 29, amp.Aux1.ToString());
            SetEditValue(hw, 30, amp.Aux2.ToString());
            SetEditValue(hw, 31, amp.Aux3.ToString());
            SetBooleanCtrlValue(hw, 10, amp.ClassicResponse);

        }


        #endregion

        #region
        //START VP SETTERS
        public static void SetVP_tune(IntPtr hw, Tune data)
        {

            SetEditValue(hw, 2, data.transpose);
            SetEditValue(hw, 3, data.Fine);
            SetEditValue(hw, 4, data.Coarse);

        }
        public static void SetVP_Chorus(IntPtr hw, Chorus data)
        {

            SetEditValue(hw, 12, data.amount);
            SetEditValue(hw, 13, data.width);
            SetEditValue(hw, 14, data.initms);

        }
        public static void SetVP_Twistaloop(IntPtr hw, Twistaloop data)
        {
            SetBooleanCtrlValue(hw, 2, data.StartAtLoop);
            SetEditValue(hw, 5, data.SpeedPrcnt);
            SetEditValue(hw, 6, data.SoopPrcnt);
        }
        public static void SetVP_Bend(IntPtr hw, Bend data)
        {
            SetEditValue(hw, 7, data.Up);
            SetEditValue(hw, 8, data.Dn);
        }

        public static void SetVP_Key(IntPtr hw, VPKey data)
        {
            SetBooleanCtrlValue(hw, 3, data.bpm);
            SetBooleanCtrlValue(hw, 4, data.latch);

            SetCBBElement(hw, 16, 5, data.assignGroup);
            SetCBBElement(hw, 17, 6, data.keymode);

            SetEditValue(hw, 9, data.delay);
            SetEditValue(hw, 10, data.offset.ToString());

        }
        public static void SetVP_Glide(IntPtr hw, Glide data)
        {

            SetEditValue(hw, 11, data.Rate.ToString());
            SetCBBElement(hw, 18, 7, data.Curve);

        }
        public static void SetVP_Env(IntPtr hw, Env data, int n)
        {
            SetBooleanCtrlValue(hw, 17 + n * 2, data.bpm);
            SetEditValue(hw, 40 + n * 12, data.Attack1Time);
            SetEditValue(hw, 46 + n * 12, data.Attack1Level);
            SetEditValue(hw, 41 + n * 12, data.Attack2Time);
            SetEditValue(hw, 47 + n * 12, data.Attack2Level);
            SetEditValue(hw, 42 + n * 12, data.Decay1Time);
            SetEditValue(hw, 48 + n * 12, data.Decay1Level);
            SetEditValue(hw, 43 + n * 12, data.Decay2Time);
            SetEditValue(hw, 49 + n * 12, data.Decay2Level);
            SetEditValue(hw, 44 + n * 12, data.Release1Time);
            SetEditValue(hw, 50 + n * 12, data.Release1Level);
            SetEditValue(hw, 45 + n * 12, data.Release2Time);
            SetEditValue(hw, 51 + n * 12, data.Release2Level);
        }
        public static void SetVP_AmpEnv(IntPtr hw, Env data)
        {
            SetVP_Env(hw, data, 0);
        }
        public static void SetVP_FilterEnv(IntPtr hw, Env data)
        {
            SetVP_Env(hw, data, 1);
        }
        public static void SetVP_AuxEnv(IntPtr hw, Env data)
        {
            SetVP_Env(hw, data, 2);
        }

        public static void SetVP_LFO1(IntPtr hw, LFO data)
        {

            SetBooleanCtrlValue(hw, 12, data.keySync);
            SetBooleanCtrlValue(hw, 11, data.bpm);
            SetEditValue(hw, 32, data.Frequency);
            SetEditValue(hw, 33, data.Delay);
            SetEditValue(hw, 34, data.Variation);
        }

        public static void SetVP_LFO2(IntPtr hw, LFO data)
        {

            SetBooleanCtrlValue(hw, 15, data.keySync);
            SetBooleanCtrlValue(hw, 14, data.bpm);
            SetEditValue(hw, 35, data.Frequency);
            SetEditValue(hw, 36, data.Delay);
            SetEditValue(hw, 37, data.Variation);

        }

        static void SetSteps(IntPtr hw, int fgnum, Step[] data)
        {
            // cannot direct query: select the stepnum first
            var bar = AutoItX.ControlGetHandle(hw, getEditElement(78 + 4 * fgnum));
            //var bar = AutoItX.ControlGetHandle(hw, ClassName_cbb_text(114 + 4 * fgnum));
            for (int i = 0; i < 64; ++i)
            {
                AutoItX.ControlSetText(hw, bar, (i + 1).ToString());
                AutoItX.ControlSend(hw, bar, "{ENTER 1}");

                SetEditValue(hw, 79 + 4 * fgnum, data[i].value.ToFring());
                SetBooleanCtrlValue(hw, 27 + 8 * fgnum, data[i].trigger);

            }
        }



        public static void SetFuncGen(IntPtr hw, int n, FuncGen data)
        {
            clickVPTab(hw, 4 + n);


            SetEditValue(hw, 76 + n * 4, data.StepRate);
            SetBooleanCtrlValue(hw, 23 + n * 8, data.bpm);
            SetCBBElement(hw, 111 + n * 8, 24 + n * 8, data.sync);
            SetBooleanCtrlValue(hw, 25 + n * 8, data.smooth);
            SetCBBElement(hw, 112 + n * 8, 26 + n * 8, data.direction);
            SetEditValue(hw, 77 + n * 4, data.endStep.ToString());
            SetSteps(hw, n, data.steps);
        }


        public static void SetCord(IntPtr hw, int n, Cord data)
        {
            SetEditValue(hw, 88 + n, data.value.ToFring());
            SetCBBElement(hw, 136 + 3 * n, 47 + 2 * n, data.source);
            SetCBBElement(hw, 138 + 3 * n, 48 + 2 * n, data.destination);

        }



        public static void SetVP_Cords(IntPtr hw, Cord[] data)
        {
            var tret = new List<Cord>();
            for (int i = 0; i < 36; ++i)
            {
                SetCord(hw, i, data[i]);
            }
        }
        #endregion


        public static void SetLinksSelection_MixTune(
            IntPtr hw,
            int linknum,
            int volume,
            int pan,
            int xpose,
            float fine
           )
        {
            SetEditValue(hw, 2 + 16 * linknum, volume.ToString());
            SetEditValue(hw, 3 + 16 * linknum, pan.ToString());
            SetEditValue(hw, 4 + 16 * linknum, xpose.ToString());
            SetEditValue(hw, 5 + 16 * linknum, fine.ToFring());
        }

        public static void SetLinksSelection_KeyWin(
            IntPtr hw,
            int linknum,
            string low,
            int fadel,
            string high,
            int fadeh
            )
        {
            SetEditValue(hw, 6 + 16 * linknum, low);
            SetEditValue(hw, 7, fadel.ToString());
            SetEditValue(hw, 8 + 16 * linknum, high);
            SetEditValue(hw, 9 + 16 * linknum, fadeh.ToString());
        }
        public static void SetLinksSelection_VelWin(
           IntPtr hw,
           int linknum,
           string low,
           int fadel,
           string high,
           int fadeh
           )
        {
            SetEditValue(hw, 10 + 16 * linknum, low);
            SetEditValue(hw, 11 + 16 * linknum, fadel.ToString());
            SetEditValue(hw, 12 + 16 * linknum, high);
            SetEditValue(hw, 13 + 16 * linknum, fadeh.ToString());
        }

        public static void SetSelectionToSplit(
           IntPtr hw,
           int linknum,
           int ccnum,
           string source,
           int low,
           int lowFade,
           int high,
           int highFade
           )
        {

            SetCBBElement(hw, 21 + linknum * 24, 2 + linknum * 2, source);

            SetEditValue(hw, 14 + 16 * linknum, low.ToString());
            SetEditValue(hw, 15 + 16 * linknum, lowFade.ToString());
            SetEditValue(hw, 16 + 16 * linknum, high.ToString());
            SetEditValue(hw, 17 + 16 * linknum, highFade.ToString());

        }

        public static void SetVoiceProcessingDatas(
            IntPtr hw,
            IntPtr ht,
            ref VoiceProcessing data)
        {
            var si = GetSelectedInstrumentNum(hw, ht);
            if (si < 0) return;
            selectNthPresetVoiceProc(hw, ht, si);
            SetVP_tune(hw, data.tune);
            SetVP_Chorus(hw, data.chorus);
            SetVP_Twistaloop(hw, data.twistaloop);
            SetVP_Bend(hw, data.bend);
            SetVP_Key(hw, data.key);
            SetVP_Glide(hw, data.glide);

            SetVP_AmpEnv(hw, data.AmpEnv);
            SetVP_FilterEnv(hw, data.FilterEnv);
            SetVP_AuxEnv(hw, data.AuxEnv);

            SetVP_LFO1(hw, data.lfo1);
            SetVP_LFO2(hw, data.lfo2);

            SetFuncGen(hw, 0, data.fg1);
            SetFuncGen(hw, 1, data.fg2);
            SetFuncGen(hw, 2, data.fg3);


            SetFilter(hw, data.filtre);
            SetAmplifier(hw, data.amp);


            SetVP_Cords(hw, data.cords);


        }

        public static VoiceProcessing GetVoiceProcessingDatas(
            IntPtr hw, IntPtr ht)
        {
            return new VoiceProcessing()
            {
                tune = getVP_tune(hw),
                chorus = getVP_Chorus(hw),
                twistaloop = getVP_Twistaloop(hw),
                bend = GetVP_Bend(hw),
                key = GetVP_Key(hw),
                glide = GetVP_Glide(hw),

                AmpEnv = GetVP_AmpEnv(hw),
                FilterEnv = GetVP_FilterEnv(hw),
                AuxEnv = GetVP_AuxEnv(hw),

                lfo1 = GetVP_LFO1(hw),
                lfo2 = GetVP_LFO1(hw),

                fg1 = GetFuncGen(hw, 0),
                fg2 = GetFuncGen(hw, 1),
                fg3 = GetFuncGen(hw, 2),

                filtre = getFilter(hw),
                amp = GetAmplifier(hw),

                cords = GetVP_Cords(hw),
            };
        }


        public static int CreateNewPreset(IntPtr hw, IntPtr ht)
        {
            AutoItX.ControlTreeView(
                  hw,
                  ht,
                   "Select",
                   "",
                   ""
               );
            AutoItX.ControlFocus(hw, ht);
            AutoItX.ControlSend(hw, ht, "^w");

            return GetSelectedInstrumentNum(hw, ht);

        }

        public static int DuplicatePresetLinks(IntPtr hw, IntPtr ht, int pnum)
        {
            var temptP = CreateNewPreset(hw, ht);

            if (selectNthPresetVoiceLinks(hw, ht, pnum) == "1")
            {
                var hc = AutoItX.ControlGetHandle(hw, ClassName_glob());
                AutoItX.ControlSend(hw, ht, "^a");
                //AutoItX.ControlSend(hw, hc, "^c");
                //coller dans un autre preset
                if (selectNthPresetVoiceLinks(hw, ht, temptP) == "1")
                {
                    var gh2 = AutoItX.ControlGetHandle(hw, ClassName_glob());
                    if ((int)gh2 > 0)
                    {
                        AutoItX.ControlClick(hw, gh2);
                        //selectAllVoices
                        AutoItX.ControlSend(hw, gh2, "^a");
                        AutoItX.ControlSend(hw, gh2, "^v");

                        var popwh = AutoItX.WinGetHandle("Paste");
                        var h = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:1]");
                        AutoItX.ControlClick(popwh, h);
                        return temptP;
                    }
                }
            }
            return -1;

        }
        public static VoiceProcessing GetDataFromSource(
            IntPtr hwn,
            IntPtr ht,
            int source)
        {
            SelectAllVoices(hwn, ht, source);
            return GetVoiceProcessingDatas(hwn, ht);
        }
        public static void CopySourceSelection(
          IntPtr hwn,
          IntPtr ht,
          int source,
          int dest)
        {
            selectNthPresetVoiceProc(hwn, ht, source);
            AutoItX.Send("^C");
            selectNthPresetVoiceProc(hwn, ht, dest);
            AutoItX.Send("^V");

        }

        static void PasteData(
           IntPtr hwn,
           IntPtr ht,
           ref VoiceProcessing data,
           int dest
           )
        {
            SelectAllVoices(hwn, ht, dest);
            SetVoiceProcessingDatas(hwn, ht, ref data);
        }



        public static void DataAutoPaste(
            IntPtr hwn,
            IntPtr ht,
            int source,
            int from,
            int to)
        {
            var dt = GetDataFromSource(hwn, ht, source);
            for (int i = from; i <= to; i++)
            {
                PasteData(hwn, ht, ref dt, i);
            }

        }



        public static void PasteVoicesTime(
            IntPtr hw,
            IntPtr hc,
            bool beforeTrue = false,
            bool newGroup = false)
        {
            var t = false;
            while (!t)
            {
                AutoItX.ControlSend(hw, hc, "^v");
                var popwh = AutoItX.WinGetHandle("Paste");
                t = (int)popwh > 0;
                if (t)
                {
                    if (beforeTrue)
                    {
                        var h = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:1]");
                        AutoItX.ControlClick(popwh, h);
                    }
                    else
                    {
                        var h = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:2]");
                        AutoItX.ControlClick(popwh, h);
                    }

                    if (newGroup)
                    {
                        var h = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:6]");
                        AutoItX.ControlClick(popwh, h);
                    }


                    var okbutt = AutoItX.ControlGetHandle(popwh, "[CLASS:Button; INSTANCE:7]");
                    AutoItX.ControlClick(popwh, okbutt);

                }


            }



        }



    }
}
