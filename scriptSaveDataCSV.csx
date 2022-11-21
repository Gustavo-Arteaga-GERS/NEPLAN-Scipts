/////////////////////////////////////////////////////////////////////////
// Author: Gustavo Arteaga
// Date:   17th November  2022
// GERS S.A.S
// scriptSaveDataCSV.csx :  NEPLAN 10 client-side csx script for exporting results in csv format. 

/////////////////////////////////////////////////////////////////////////

using BCP.Neplan.Helpers.Script;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;



ShowMessageWindow();
ClearMessages();
SendInfoMessage("***************************************************************************************** ");
SendInfoMessage("************* Start of the script to print the location of the contingencies ************ ");
SendInfoMessage("***************************************************************************************** ");


// Enter your folder path here
string path = "G:\\My Drive\\STN\\Scripts\\ScriptSaveData";
////

Dictionary<string, List<string>> AllAnalysisResults = new Dictionary<string, List<string>>();
List<string> listaDeDiagramas = new List<string>();
listaDeDiagramas = GetDiagrams();



Dictionary<string, Dictionary<string, List<string>>> results = new Dictionary<string, Dictionary<string, List<string>>>();


Dictionary<string, string> allElementOfVariante = new Dictionary<string, string>();
GetAllElementsOfVariant(out allElementOfVariante);
SendInfoMessage("-- > All elements of the variant were obtained");


string inputElementsContingency = @"" + path + "\\elementsContingency.csv";
string inputElementsForAnalysis = @"" + path + "\\elementsForAnalysis.csv";


string[] elementsContingency = File.ReadAllLines(inputElementsContingency);
string[] elementsAnalysis = File.ReadAllLines(inputElementsForAnalysis);



//Foreach in all elements obtained in the varian
foreach (var element in allElementOfVariante)
{
    //SendInfoMessage(element.Key);
    //Foreach in all contingencies input
    foreach (var contingencies in elementsContingency)
    {
        var contingency = contingencies.Split(',');
        //SendInfoMessage(contingency[0]);
        if (element.Key == contingency[0])
        {
            //SendInfoMessage("👍👍");
            bool openPort0 = SwitchElementAtPort(element.Key, 0, false);
            bool openPort1 = SwitchElementAtPort(element.Key, 1, false);
            if (openPort0 == true)
            {
                SendInfoMessage("-- > Element " + element.Key + " open");
                SendInfoMessage("-- > Run Current Analysis ");
                bool coreranalisis = RunCurrentAnalysis();
                SendInfoMessage("-- > Update Diagram ");
                UpdateTopologyAndDiagram();
                AllAnalysisResults = GetAllElementResults();
                Dictionary<string, List<string>> resultsElement = new Dictionary<string, List<string>>();
                foreach (string key in AllAnalysisResults.Keys)
                {
                    //Foreach in all celements for Analysis input
                    foreach (var e in elementsAnalysis)
                    {
                        var elementAnalysis = e.Split(',');

                        if (elementAnalysis[0] == key)
                        {
                            SendInfoMessage("Result of element  " + key);
                            
                            List<string> resultsNEPLAN = AllAnalysisResults[key];
                            List<string> resultForPrint = new List<string>();
                            resultForPrint.Add(elementAnalysis[1]);

                            if (elementAnalysis[1] == "line")
                            {
                                
                                SendInfoMessage("Line Resuls:");
                                SendInfoMessage(resultsNEPLAN[23]);
                                SendInfoMessage(resultsNEPLAN[24]);
                                SendInfoMessage(resultsNEPLAN[26]);
                                SendInfoMessage(resultsNEPLAN[29]);

                                resultForPrint.Add(resultsNEPLAN[23].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[24].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[26].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[29].Split(':')[1]);
                                
                            }
                            if (elementAnalysis[1] == "node")
                            {
                                SendInfoMessage("Node Results:");
                                SendInfoMessage(resultsNEPLAN[0]);
                                SendInfoMessage(resultsNEPLAN[1]);

                                resultForPrint.Add(resultsNEPLAN[0].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[1].Split(':')[1]);
                            }
                            if (elementAnalysis[1] == "trafo2")
                            {
                                SendInfoMessage("Transformer 2w Results:");
                                SendInfoMessage("Port 0");
                                SendInfoMessage(resultsNEPLAN[20]);
                                SendInfoMessage(resultsNEPLAN[21]);
                                SendInfoMessage(resultsNEPLAN[26]);

                                SendInfoMessage("Port 1");
                                SendInfoMessage(resultsNEPLAN[33]);
                                SendInfoMessage(resultsNEPLAN[34]);

                                resultForPrint.Add(resultsNEPLAN[20].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[21].Split(':')[1]);
                                //resultForPrint.Add(resultsNEPLAN[26].Split(':')[1]);

                                resultForPrint.Add(resultsNEPLAN[33].Split(':')[1]);
                                resultForPrint.Add(resultsNEPLAN[34].Split(':')[1]);
                            }
                            //SendInfoMessage("_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _");
                            resultsElement.Add(key, resultForPrint);
                          
                        }
                    }
                }

                results.Add(element.Key, resultsElement);
              
            }

            SendInfoMessage("-- > close element" + element.Key);
            openPort0 = SwitchElementAtPort(element.Key, 0, true);
            openPort1 = SwitchElementAtPort(element.Key, 1, true);
            SendInfoMessage("-- > element closed" + element.Key);
        }
    }
}





//Create new CSV with results
string dateNow = DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss");
string fileName = @"" + path + "\\results\\" + dateNow + ".csv";
using (StreamWriter fileResults = File.CreateText(fileName))
{
    fileResults.WriteLine("Contingency,ElementName,TypeElement,P,Q,I,Loading");
    foreach (var itemResults in results)
    {
        //SendInfoMessage(itemResults.Value);
        foreach (var i in elementsAnalysis)
        {
            var itemElement = i.Split(',');
            foreach (var itemElemets in itemResults.Value)
            {
                if (itemElement[0] == itemElemets.Key)
                {
                    var lineWrite = "";
                    foreach (var item in itemElemets.Value)
                    {

                        lineWrite = lineWrite + "," + item;

                    }
                    fileResults.WriteLine(itemResults.Key + "," + itemElemets.Key  + lineWrite);
                }
            }
        }

    }
}


UpdateTopologyAndDiagram();
SendInfoMessage("-- > Diagrams generated successfully.");
SendInfoMessage("-- > End script (csx) 👍");


