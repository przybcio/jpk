using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using JpkApp.Model;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JpkApp
{
    internal class Program
    {
        private const int StartRowIndex = 9;

        private static void Main(string[] args)
        {
            string GetSheetName()
            {
                return $"Kwartał {args[2]}-{args[1]}";
            }

            string GetFilenameFromDate()
            {
                var filename = DateTime.Now.AddMonths(-1).ToString("MMMM-yyyy");
                return filename + ".xml";
            }

            JPK InitJpk()
            {
                var dateFrom = DateTime.Parse($"01-{args[0]}-{args[1]}");
                var dateTo = dateFrom.AddMonths(1).AddDays(-1);

                return new JPK
                {
                    Naglowek = new TNaglowek
                    {
                        KodFormularza = new TNaglowekKodFormularza {Value = TKodFormularza.JPK_VAT},
                        DataOd = dateFrom,
                        DataDo = dateTo,
                        DataWytworzeniaJPK = DateTime.Now,
                        NazwaSystemu = "jpkApp"
                        //WariantFormularza = ??
                    },
                    Podmiot1 = new JPKPodmiot1
                    {
                        Email = "lukas.przybyl@gmail.com",
                        NIP = "8361774141",
                        PelnaNazwa = "Idev Łukasz Przybył"
                    },
                    ZakupCtrl = new JPKZakupCtrl(),
                    SprzedazCtrl = new JPKSprzedazCtrl()
                };
            }

            void ProcessZakup(int zakupRowNum, ISheet zakupSheet, JPK jpk)
            {
                var zakupIndex = 0;
                var lpZakupu = 1;
                for (var row = StartRowIndex; row < zakupRowNum; row++, zakupIndex++)
                {
                    var issueDate = zakupSheet.GetRow(row)?.GetCell(2).DateCellValue ?? DateTime.Now;
                    if (issueDate <= jpk.Naglowek.DataDo && issueDate >= jpk.Naglowek.DataOd)
                    {

                        var jpkZakupWiersz =
                            new JPKZakupWiersz
                            {
                                AdresDostawcy = zakupSheet.GetRow(row)?.GetCell(7).StringCellValue,
                                DataWplywu = issueDate,
                                DataWplywuSpecified = true,
                                DataZakupu = issueDate,
                                LpZakupu = lpZakupu++.ToString(),
                                NazwaDostawcy = zakupSheet.GetRow(row)?.GetCell(6).StringCellValue,
                                NrDostawcy =
                                    zakupSheet.GetRow(row)?.GetCell(5).NumericCellValue
                                        .ToString(CultureInfo.InvariantCulture),
                                K_45 = Convert.ToDecimal(zakupSheet.GetRow(row)?.GetCell(9).NumericCellValue),
                                K_46 = Convert.ToDecimal(zakupSheet.GetRow(row)?.GetCell(15).NumericCellValue)
                            };
                        var cell = zakupSheet.GetRow(row)?.GetCell(1);
                        if (cell != null)
                            switch (cell.CellType)
                            {
                                case CellType.Unknown:
                                    break;
                                case CellType.Numeric:
                                    jpkZakupWiersz.DowodZakupu =
                                        cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                                    break;
                                case CellType.String:
                                    jpkZakupWiersz.DowodZakupu = cell.StringCellValue;
                                    break;
                                case CellType.Formula:
                                    break;
                                case CellType.Blank:
                                    break;
                                case CellType.Boolean:
                                    break;
                                case CellType.Error:
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException();
                            }
                        jpk.ZakupWiersz[zakupIndex] = jpkZakupWiersz;
                    }
                }

                jpk.ZakupCtrl.LiczbaWierszyZakupow = jpk.ZakupWiersz.Count(wiersz => wiersz != null).ToString();
                jpk.ZakupCtrl.PodatekNaliczony = jpk.ZakupWiersz.Where(wiersz => wiersz != null).Sum(wiersz => wiersz.K_46);
            }

            void ProcessSprzedaz(int zakupRowNum, ISheet sprzedazSheet, JPK jpk)
            {
                var sprzedazIndex = 0;
                var lpSprzedazy = 1;
                for (var row = StartRowIndex; row < zakupRowNum; row++, sprzedazIndex++)
                {
                    var issueDate = sprzedazSheet.GetRow(row)?.GetCell(2).DateCellValue ?? DateTime.Now;
                    if (issueDate <= jpk.Naglowek.DataDo && issueDate >= jpk.Naglowek.DataOd)
                    {
                        var jpkSprzedazWiersz =
                            new JPKSprzedazWiersz
                            {
                                AdresKontrahenta = sprzedazSheet.GetRow(row)?.GetCell(5).StringCellValue,
                                DataSprzedazy = sprzedazSheet.GetRow(row)?.GetCell(3).DateCellValue ?? DateTime.Now,
                                DataSprzedazySpecified = true,
                                DataWystawienia = issueDate,
                                LpSprzedazy = lpSprzedazy++.ToString(),
                                NazwaKontrahenta = sprzedazSheet.GetRow(row)?.GetCell(5).StringCellValue,
                                NrKontrahenta = sprzedazSheet.GetRow(row)?.GetCell(4).NumericCellValue
                                    .ToString(CultureInfo.InvariantCulture),
                                K_19 = Convert.ToDecimal(sprzedazSheet.GetRow(row)?.GetCell(8).NumericCellValue),
                                K_20 = Convert.ToDecimal(sprzedazSheet.GetRow(row)?.GetCell(14).NumericCellValue)
                            };
                        var cell = sprzedazSheet.GetRow(row)?.GetCell(1);
                        if (cell != null)
                            switch (cell.CellType)
                            {
                                case CellType.Unknown:
                                    break;
                                case CellType.Numeric:
                                    jpkSprzedazWiersz.DowodSprzedazy =
                                        cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                                    break;
                                case CellType.String:
                                    jpkSprzedazWiersz.DowodSprzedazy = cell.StringCellValue;
                                    break;
                                case CellType.Formula:
                                    break;
                                case CellType.Blank:
                                    break;
                                case CellType.Boolean:
                                    break;
                                case CellType.Error:
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException();
                            }
                        jpk.SprzedazWiersz[sprzedazIndex] = jpkSprzedazWiersz;
                    }
                }

                jpk.SprzedazCtrl.LiczbaWierszySprzedazy = jpk.SprzedazWiersz.Count(wiersz => wiersz != null).ToString();
                jpk.SprzedazCtrl.PodatekNalezny = jpk.SprzedazWiersz.Where(wiersz => wiersz != null).Sum(wiersz => wiersz.K_20);
            }

            int GetSprzedazRowNum(ISheet sheet)
            {
                return GetZakupRowNum(sheet);
            }

            int GetZakupRowNum(ISheet sheet)
            {
                var index = StartRowIndex + 1;
                while (Math.Abs(sheet.GetRow(index).GetCell(0).NumericCellValue - default(double)) > 0.1)
                    index++;
                return index;
            }

            var initJpk = InitJpk();

            using (var fsZakup = new FileStream(@"C:\tmp\Ewidencja nabycia towarów i usług.xls.xlsx", FileMode.Open,
                FileAccess.Read))
            {
                

                var xssfWorkbook = new XSSFWorkbook(fsZakup);
                var zakupSheet = xssfWorkbook.GetSheet(GetSheetName());
                var zakupRowNum = GetZakupRowNum(zakupSheet);
                initJpk.InitZakupArray(zakupRowNum - StartRowIndex);

                using (var fsSprzedaz = new FileStream(@"C:\tmp\Ewidencja sprzedaży.xls.xlsx",
                    FileMode.Open,
                    FileAccess.Read))
                {
                    var xssfSprzedazWorkbook = new XSSFWorkbook(fsSprzedaz);
                    var sprzedazSheet = xssfSprzedazWorkbook.GetSheet(GetSheetName());
                    var sprzedazRowNum = GetSprzedazRowNum(sprzedazSheet);
                    initJpk.InitSprzedazArray(sprzedazRowNum - StartRowIndex);

                    ProcessSprzedaz(sprzedazRowNum, sprzedazSheet, initJpk);
                }

                ProcessZakup(zakupRowNum, zakupSheet, initJpk);

                using (var fileStreamWrite = new FileStream(GetFilenameFromDate(), FileMode.Create, FileAccess.Write))
                {
                    var xmlSerializer = new XmlSerializer(typeof(JPK));
                    xmlSerializer.Serialize(fileStreamWrite, initJpk);
                }
            }
        }
    }
}