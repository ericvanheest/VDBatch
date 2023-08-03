using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace VDBatch
{
    public class CLOptions
    {
        enum NextOption { JobFile, OutputFile, None };
        enum NextArgs { None, InputFile, Last };

        public string Usage { get { return GetUsage(); } }
        public bool Help { get; set; }
        public List<string> Warnings { get; set; }
        public List<string> Errors { get; set; }

        public bool ExitAfterOutput { get; set; }
        public string UseJobFile { get; set; }
        public string UseOutputFile { get; set; }

        public List<string> InputFiles{ get; set; }

        static string Version()
        {
            return System.Diagnostics.Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductVersion;
        }

        public static string GetUsage()
        {
            return string.Format("VDBatch, version {0}\r\n\r\n" +
            "Usage:    VDBatch.exe [options] [files]\r\n\r\n" +
            "Options:\r\n" +
            "  -j [file]\tLoads the specified file as the template\r\n" +
            "  -c [file]\tCreates the specified file from the current template and file list\r\n" +
            "  -e\tExit the program after creating an output file\r\n" +
            "\r\n" +
            "Any other arguments are interpreted as files to add to the processing list.\r\n" +
            "\r\n" +
            "Example:\r\n" +
            "\r\n" +
            "VDBatch.exe - j MyJobTemplate.job - c Virtualdub.jobs file1.avi vid*.avi\r\n" +
            ""
            , Version());
        }

        public bool HasErrors { get { return Errors.Count > 0; } }
        public bool HasWarnings { get { return Warnings.Count > 0; } }

        private long GetLong(string strValue)
        {
            long lVal;
            if (Int64.TryParse(strValue, out lVal))
                return lVal;

            if (!strValue.StartsWith("0x") && !strValue.StartsWith("-0x"))
            {
                Warnings.Add(String.Format("Could not interpret \"{0}\" as an integer.", strValue));
                return 0;
            }

            int iStart = 2;
            int iMultiply = 1;

            if (strValue[0] == '-')
            {
                iStart = 3;
                iMultiply = -1;
            }

            try
            {
                return Convert.ToInt64(strValue.Substring(iStart), 16) * iMultiply;
            }
            catch (Exception ex)
            {
                Warnings.Add(String.Format("Could not interpret \"{0}\" as an integer, Exception: {1}", strValue, ex.Message));
            }
            return 0;
        }

        public CLOptions(string[] args, bool bSkipZeroArg = false)
        {
            Help = false;
            Errors = new List<string>();
            Warnings = new List<string>();

            InputFiles = new List<string>();
            ExitAfterOutput = false;
            UseJobFile = "";
            UseOutputFile = "";

            NextOption nextOption = NextOption.None;
            NextArgs nextArg = NextArgs.InputFile;

            foreach (string strArg in args)
            {
                if (bSkipZeroArg)
                {
                    bSkipZeroArg = false;
                    continue;
                }

                switch (nextOption)
                {
                    case NextOption.None:
                        if (strArg[0] == '-' || strArg[0] == '/')
                        {
                            // Arguments passed on the command line with a leading hyphen or slash
                            for (int i = 1; i < strArg.Length; i++)
                            {
                                switch (strArg[i])
                                {
                                    case 'h':
                                    case 'H':
                                    case '?':
                                        Help = true;
                                        break;

                                    case 'c':
                                    case 'C':
                                        nextOption = NextOption.OutputFile;
                                        break;

                                    case 'e':
                                    case 'E':
                                        ExitAfterOutput = true;
                                        break;

                                    case 'j':
                                    case 'J':
                                        nextOption = NextOption.JobFile;
                                        break;

                                    case '-':
                                        // Options starting with --
                                        switch (strArg.Substring(i + 1))
                                        {
                                            case "help":
                                                Help = true;
                                                break;
                                            default:
                                                Warnings.Add(String.Format("Ignoring unknown option '{0}'", strArg.Substring(i + 1)));
                                                break;
                                        }
                                        i = strArg.Length;
                                        break;
                                    default:
                                        Warnings.Add(String.Format("Ignoring unknown option '{0}'", strArg[i]));
                                        break;
                                }
                            }
                        }
                        else
                        {
                            // Arguments passed on the command line without a leading hyphen
                            switch (nextArg)
                            {
                                case NextArgs.InputFile:
                                    InputFiles.Add(strArg);
                                    // Don't increment nextArg; all arguments are input files
                                    break;
                                default:
                                    Warnings.Add(String.Format("Ignoring extra command line argument '{0}'", strArg));
                                    break;
                            }
                        }
                        break;

                    case NextOption.JobFile:
                        UseJobFile = strArg;
                        nextOption = NextOption.None;
                        break;

                    case NextOption.OutputFile:
                        UseOutputFile = strArg;
                        nextOption = NextOption.None;
                        break;

                    default:
                        Errors.Add("Unknown next option");
                        break;
                }
            }
        }
    }
}
