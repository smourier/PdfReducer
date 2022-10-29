using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Acrobat;

namespace PdfReducer
{
    class Program
    {
        private static bool alwaysRewrite;
        private static bool failIfExists;
        private static bool dontCopyOnError;
        private static bool skipExisting;

        static void Main()
        {
            if (Debugger.IsAttached)
            {
                RealMain();
            }
            else
            {
                try
                {
                    RealMain();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        static void Help()
        {
            Console.WriteLine("Format:");
            Console.WriteLine();
            Console.WriteLine(Assembly.GetEntryAssembly().GetName().Name.ToUpperInvariant() + " <input file or directory path> <output file or directory path> [options]");
            Console.WriteLine();
            Console.WriteLine("Description:");
            Console.WriteLine("    This tool is used to reduce PDF files size using Adobe Acrobat.");
            Console.WriteLine("    Input and output paths must be different.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("    /failIfExists    If set, existing file(s) will not be overwritten and an error will be raised.");
            Console.WriteLine("    /alwaysRewrite   If set, always rewrite output file(s) even if its size is bigger.");
            Console.WriteLine("    /dontCopyOnError If set, do not copy files that cannot be reduced because an error occurred.");
            Console.WriteLine("    /skipExisting    If set, skip existing target files.");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine();
            Console.WriteLine("    " + Assembly.GetEntryAssembly().GetName().Name.ToUpperInvariant() + " input.pdf output.pdf /alwaysrewrite");
            Console.WriteLine();
            Console.WriteLine("    Opens input.pdf and saves it back to output.pdf.");
            Console.WriteLine("    If output.pdf already exists, it will be replaced without any warning.");
            Console.WriteLine("    If output.pdf size is bigger than input.pdf, it will be kept.");
            Console.WriteLine();
        }

        static void RealMain()
        {
            Console.WriteLine("PdfReducer - Copyright (C) 2021-" + DateTime.Now.Year + " Simon Mourier. All rights reserved.");
            Console.WriteLine();

            var inputPath = CommandLine.GetNullifiedArgument(0);
            var outputPath = CommandLine.GetNullifiedArgument(1);
            if (outputPath == null)
            {
                outputPath = ".";
            }

            if (CommandLine.HelpRequested || inputPath == null)
            {
                Help();
                return;
            }

            inputPath = Path.GetFullPath(inputPath);
            outputPath = Path.GetFullPath(outputPath);
            if (inputPath.EqualsIgnoreCase(outputPath))
            {
                Help();
                return;
            }

            failIfExists = CommandLine.GetArgument<bool>(nameof(failIfExists));
            alwaysRewrite = CommandLine.GetArgument<bool>(nameof(alwaysRewrite));
            dontCopyOnError = CommandLine.GetArgument<bool>(nameof(dontCopyOnError));
            skipExisting = CommandLine.GetArgument<bool>(nameof(skipExisting));
            Console.WriteLine("Input Path          : " + inputPath);
            Console.WriteLine("Output Path         : " + outputPath);
            Console.WriteLine("Fail If Exists      : " + failIfExists);
            Console.WriteLine("Always Rewrite      : " + alwaysRewrite);
            Console.WriteLine("Don't Copy On Error : " + dontCopyOnError);
            Console.WriteLine("Skip Existing       : " + skipExisting);
            Console.WriteLine();

            var app = new AcroApp();
            app.Hide();
            app.CloseAllDocs();

            long oldSize = 0;
            long newSize = 0;
            if (Directory.Exists(inputPath))
            {
                var tuple = ReduceDirectory(app, inputPath, outputPath);
                oldSize += tuple.Item1;
                newSize += tuple.Item2;
            }
            else
            {
                // assume input is file
                if (Directory.Exists(outputPath))
                {
                    var tuple = ReduceFile(app, inputPath, Path.Combine(outputPath, Path.GetFileName(inputPath)));
                    oldSize += tuple.Item1;
                    newSize += tuple.Item2;
                }
                else
                {
                    // assume output is file
                    var tuple = ReduceFile(app, inputPath, outputPath);
                    oldSize += tuple.Item1;
                    newSize += tuple.Item2;
                }
            }

            app.CloseAllDocs();
            app.Exit();

            Console.WriteLine();
            Console.WriteLine("Old Bytes           : " + oldSize + " (" + Conversions.FormatByteSize(oldSize) + ")");
            Console.WriteLine("New Bytes           : " + newSize + " => " + Conversions.FormatByteSize(newSize) + ")");

            var savedBytes = oldSize - newSize;
            if (savedBytes >= 0)
            {
                Console.WriteLine("Saved Bytes         : " + savedBytes + " (" + Conversions.FormatByteSize(savedBytes) + ")");
                Console.WriteLine("Saved Percent       : " + (100f * (oldSize - newSize) / oldSize) + " %");
            }
            else
            {
                Console.WriteLine("Lost Bytes          : " + savedBytes + " (" + Conversions.FormatByteSize(savedBytes) + ")");
                Console.WriteLine("Lost Percent        : " + (100f * (newSize - oldSize) / newSize) + " %");
            }
        }

        static Tuple<long, long> ReduceDirectory(AcroApp app, string sourcePath, string targetPath)
        {
            long oldSize = 0;
            long newSize = 0;
            foreach (var file in Directory.EnumerateFiles(sourcePath, "*.pdf"))
            {
                var targetFilePath = Path.Combine(targetPath, Path.GetFileName(file));
                var tuple = ReduceFile(app, file, targetFilePath);
                oldSize += tuple.Item1;
                newSize += tuple.Item2;
            }

            foreach (var dir in Directory.EnumerateDirectories(sourcePath))
            {
                var targetDir = Path.Combine(targetPath, Path.GetFileName(dir));
                var tuple = ReduceDirectory(app, dir, targetDir);
                oldSize += tuple.Item1;
                newSize += tuple.Item2;
            }

            return new Tuple<long, long>(oldSize, newSize);
        }

        static Tuple<long, long> ReduceFile(AcroApp app, string sourcePath, string targetPath)
        {
            var oldSize = new FileInfo(sourcePath).Length;
            if (skipExisting && File.Exists(targetPath) && new FileInfo(targetPath).Length != 0)
            {
                Console.WriteLine("'" + sourcePath + "' => '" + targetPath + "' (skipped because it already exists)");
                return new Tuple<long, long>(oldSize, oldSize);
            }

            // open doc
            var doc = new AcroAVDoc();
            try
            {
                doc.Open(sourcePath, sourcePath);

                // extract page into new (temp) doc
                var source = (AcroPDDoc)doc.GetPDDoc();
                try
                {
                    object jso = source.GetJSObject(); // can't use C# dynamic with Acrobat automation

                    // https://opensource.adobe.com/dc-acrobat-sdk-docs/acrobatsdk/pdfs/acrobatsdk_jsapiref.pdf
                    var target = jso.InvokeMember("extractPages");

                    if (File.Exists(targetPath))
                    {
                        if (failIfExists)
                            throw new Exception("File '" + targetPath + "' already exists.");

                        File.Delete(targetPath);
                    }
                    else
                    {
                        Extensions.EnsureFileDirectoryExists(targetPath);
                    }

                    // save new doc
                    app.TrustFilePath(targetPath);
                    target.InvokeMember("saveAs", targetPath);
                    target.InvokeMember("closeDoc");
                }
                finally
                {
                    source.Close();
                    doc.Close(1);
                    app.CloseAllDocs();
                }
            }
            catch (Exception ex)
            {
                if (!dontCopyOnError)
                {
                    Extensions.EnsureFileDirectoryExists(targetPath);
                    File.Copy(sourcePath, targetPath, true);
                    Console.WriteLine("'" + sourcePath + "' => '" + targetPath + "' (copied as is after error: " + ex.GetAllMessagesWithDots() + ")");
                }
                else
                {
                    Console.WriteLine("'" + sourcePath + "' (skipped after error: " + ex.GetAllMessagesWithDots() + ")");
                }
                return new Tuple<long, long>(oldSize, oldSize);
            }

            var newSize = new FileInfo(targetPath).Length;
            if (newSize > oldSize)
            {
                // keep original
                if (!alwaysRewrite)
                {
                    File.Copy(sourcePath, targetPath, true);
                    Console.WriteLine("'" + sourcePath + "' => '" + targetPath + "' (not reduced, was kept as is)");
                }
                else
                {
                    Console.WriteLine("'" + sourcePath + "' => '" + targetPath + "' (result size is bigger)");
                }
                return new Tuple<long, long>(oldSize, oldSize);
            }
            Console.WriteLine("'" + sourcePath + "' => '" + targetPath + "'");
            return new Tuple<long, long>(oldSize, newSize);
        }
    }
}
