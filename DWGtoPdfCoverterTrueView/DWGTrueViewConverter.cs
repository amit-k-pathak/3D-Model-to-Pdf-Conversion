using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;
using AccoreconsoleLauncher;
using System.Text.RegularExpressions;

namespace DWGtoPdfCoverterTrueView
{
    class DWGTrueViewConverter
    {
        private BackgroundWorker _worker;
        private Stopwatch _watch;
        private string _fileName;
        List<string> _colA;
        List<string> _colC;
        BathConversionStatus _st;
        private int _successCount;
        private int _fileCount;
        private int _failedCount;
        private long _avgTime;

        public DWGTrueViewConverter(string csvFile)
        {
            this._worker = null;
            this._watch = null;
            this._fileName = csvFile;
            this._st = BathConversionStatus.Failed;
            this._successCount = 0;
            this._fileCount = 0;
            this._failedCount = 0;
            this._avgTime = 0;
        }

        #region worker

        private void Init()
        {
            if (this._worker == null)
            {
                this._worker = new BackgroundWorker();
                this._worker.WorkerReportsProgress = true;
                this._worker.WorkerSupportsCancellation = false;
                this._worker.DoWork += new DoWorkEventHandler(HandleDoWork);
                this._worker.ProgressChanged += new ProgressChangedEventHandler(HandleProgressChanged);
                this._worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(HandleWorkerCompleted);
                this._watch = new Stopwatch();
            }
        }

        private void StartWorker(Accoreconsole console)
        {
            Init();

            if (!_worker.IsBusy)
            {
                _watch.Start();
                this._worker.RunWorkerAsync(console);
            }
        }

        private void HandleDoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Accoreconsole console = e.Argument as Accoreconsole;

            int i = 0;
            string msg = "";

            if (!(worker == null))
            {
                if (!(console == null))
                {
                    for (int index = 0; index < _colA.Count; ++index)
                    {
                        if (console.ConvertDwgToPdf(_colC[index]) == 0)
                            ++i;
                        msg = console.GetLog();
                        worker.ReportProgress(i, msg);
                    }
                }
            }
        }

        private void HandleWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (!(worker == null))
            {
                if (e.Error == null)
                {
                    Console.WriteLine("Seems like, all done !!");
                    _watch.Stop();
                    Console.WriteLine("Total Time : " + _watch.ElapsedMilliseconds.ToString());
                }
                else
                    Console.WriteLine("failed : " + e.Error.Data);
            }
        }

        private void HandleProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (!(worker == null))
            {
                string msg = e.UserState as string;

                if (!string.IsNullOrWhiteSpace(msg))
                {
                    string[] str = msg.Split('\n');
                    foreach (string s in str)
                        Console.WriteLine(s);

                    if (e.ProgressPercentage == _successCount + 1)
                    {
                        _successCount++;
                    }
                    else
                    {
                        _failedCount++;
                    }
                    _watch.Stop();
                    Console.WriteLine("Total Time : " + _watch.ElapsedMilliseconds.ToString());
                    Console.WriteLine("Converted Files = " + _successCount + " In Process : " + (_fileCount - (_successCount + _failedCount)) + " Failed = " + _failedCount);
                }
            }
        }

        #endregion

        public Status ReadCsvFile()
        {
            Status readStatus = Status.FileReadFailed;
            StreamReader reader = null;

            if (!File.Exists(this._fileName))
                return Status.FileNotFound;

            try
            {
                reader = new StreamReader(File.OpenRead(this._fileName));
                this._colA = new List<string>();
                this._colC = new List<string>();
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        string[] values = line.Split(',');
                        //values[0].Replace(@"\\", @"\");

                        if (values.Count() != 3)
                            return Status.FileReadFailed;

                        this._colA.Add(values[0]);

                        //Console.WriteLine(values[2]);
                        //Regex.Replace(values[2], @"\\\\", @"/");

                        if (!string.IsNullOrWhiteSpace(values[2]))
                        {
                            int counter = 0;
                            char[] parse = values[2].ToCharArray();

                            for (int i = 0; i < parse.Count(); ++i)
                            {
                                if (parse[i] == '\\')
                                {
                                    parse[counter++] = '\\';
                                    while (parse[i] == '\\')
                                        i++;
                                    i--;
                                }
                                else if (parse[i] == '>')
                                {
                                    if ((i > 0 && parse[i - 1] == ' ') || ((i + 1) < parse.Count() && parse[i + 1] == ' '))
                                    {
                                        if (parse[i + 1] == ' ')
                                            parse[counter++] = '\\';
                                        else
                                            counter--;
                                    }
                                    continue;
                                }
                                else
                                {
                                    if (i > 0 && parse[i - 1] == '>')
                                        continue;
                                    parse[counter++] = parse[i];
                                }
                            }

                            //Console.WriteLine(parse);
                            if(parse.Count() >= 2 && parse[1] != ':')
                                this._colC.Add(Path.GetDirectoryName(this._fileName) + "\\" + new string(parse, 0, counter));
                            else
                                this._colC.Add(new string(parse, 0, counter));
                        }
                    }
                }
                this._fileCount = this._colC.Count;
                readStatus = Status.FileReadSuccess;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\n" + ex.StackTrace);
                readStatus = Status.FileReadFailed;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Dispose();
                    reader = null;
                }
            }

            return readStatus;
        }

        public void Convert()
        {
            //Accoreconsole console = new Accoreconsole();
            var newLine = string.Format("{0},{1}", "Pdf File Name", "Time(ms)");
            var csv = new StringBuilder();
            csv.AppendLine(newLine);   
            try
            {
                Console.WriteLine("Reading csv file " + this._fileName);
                Status st = ReadCsvFile();
                if (st == Status.FileReadSuccess)
                    Console.WriteLine("Success..."+this._fileName);
                else if (st == Status.FileNotFound)
                    Console.WriteLine("File not found..." + this._fileName);
                else
                {
                    Console.WriteLine("File reading failed..." + this._fileName);
                    return;
                }
                //StartWorker(console);  //for BG Worker, not needed 
                if (this._watch == null)
                    this._watch = new Stopwatch();

                for (int i = 0; i < this._colC.Count; ++i)
                {
                    Accoreconsole console = new Accoreconsole();
                    Console.WriteLine("Converting file : " + _colC[i] + "...");
                    
                    if (i == 0)
                        _watch.Start();
                    else
                        _watch.Restart();

                    if (console.ConvertDwgToPdf(_colC[i]) == 0)
                    {
                        _successCount++;
                        Console.WriteLine("Success : " + _colC[i]);
                    }
                    else
                        _failedCount++;
                    _watch.Stop();
                    newLine = string.Format("{0},{1}", Path.GetDirectoryName(_colC[i])+"\\"+Path.GetFileNameWithoutExtension(_colC[i])+"_Conv.pdf", _watch.ElapsedMilliseconds.ToString());
                    csv.AppendLine(newLine); 
                    string[] str = console.GetLog().Split('\n');
                    foreach (string s in str)
                        Console.WriteLine(s);
                    this._avgTime += _watch.ElapsedMilliseconds;
                    Console.WriteLine("Total Time : " + _watch.ElapsedMilliseconds.ToString() + " ms");
                    Console.WriteLine("Total Files = " + _fileCount + " Success = " + _successCount + " In Process : " + (_fileCount - (_successCount + _failedCount)) + " Failed = " + _failedCount);
                    Console.WriteLine("---------------------------------------------------------------");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception : " + ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                if (_watch != null)
                {
                    _watch.Stop();
                    Console.WriteLine("Total Time : " + _watch.ElapsedMilliseconds.ToString() + " ms");
                    _watch = null;
                }
               
                csv.AppendLine();
                csv.AppendLine(String.Format("{0},{1}", "Total File Count", this._fileCount));
                csv.AppendLine(String.Format("{0},{1}", "Success", this._successCount));
                csv.AppendLine(String.Format("{0},{1}", "Failed", this._failedCount));
                csv.AppendLine(String.Format("{0},{1}", "Success Rate", (this._successCount / (this._fileCount * 1.0)) * 100 + " %"));
                csv.AppendLine(String.Format("{0},{1}", "Average Time", (this._avgTime*1.0)/(this._successCount + this._failedCount)));
                File.WriteAllText("csv_log.csv", csv.ToString());
                csv = null;
            }
        }

        public void Dispose()
        {
            if (this._colA != null)
                this._colA = null;

            if (this._colC != null)
                this._colC = null;
        }
   }
}
