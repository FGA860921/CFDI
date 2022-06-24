// Decompiled with JetBrains decompiler
// Type: ReadCFDIfromMail.Program
// Assembly: ReadCFDIfromMail, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 5C3D3298-27B9-4099-8001-CBF5C018B0E0
// Assembly location: C:\tmp\ReadCFDIfromMail\ReadCFDIfromMail.dll

using Ionic.Zip;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using MimeKit.Tnef;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;

namespace ReadCFDIfromMail
{
  internal class Program
  {
    private static void Main(string[] args)
    {
      List<CfdiHeaderVO> cfdiHeaderVoList = new List<CfdiHeaderVO>();
      string appSetting1 = ConfigurationManager.AppSettings["OutputPath"];
      string str = DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString();
      string path1 = appSetting1 + "CFDI_DATA_" + str + ".txt";
      string path2 = appSetting1 + "CFDI_RELA_" + str + ".txt";
      using (ImapClient imapClient = new ImapClient())
      {
        string appSetting2 = ConfigurationManager.AppSettings["ImapServer"];
        int num1 = int.Parse(ConfigurationManager.AppSettings["ImapPort"]);
        bool flag = bool.Parse(ConfigurationManager.AppSettings["ImapSSL"]);
        imapClient.Connect(appSetting2, num1, flag, new CancellationToken());
        ((MailService) imapClient).AuthenticationMechanisms.Remove("XOAUTH2");
        imapClient.Authenticate(ConfigurationManager.AppSettings["ImapUser"], ConfigurationManager.AppSettings["ImapPassword"], new CancellationToken());
        int num2 = int.Parse(ConfigurationManager.AppSettings["DaysToRead"]);
        BinarySearchQuery binarySearchQuery = SearchQuery.Not((SearchQuery) SearchQuery.HasGMailLabel("PROCESSED")).And((SearchQuery) SearchQuery.Not((SearchQuery) SearchQuery.HasGMailLabel("NoXmlFound"))).And((SearchQuery) SearchQuery.DeliveredAfter(DateTime.Now.AddDays((double) -num2)));
        IMailFolder folder1 = imapClient.GetFolder(imapClient.PersonalNamespaces[0]);
        List<string> stringList1 = new List<string>();
        stringList1.Add("PROCESSED");
        List<string> stringList2 = new List<string>();
        stringList2.Add("NoXmlFound");
        CancellationToken cancellationToken = new CancellationToken();
        foreach (IMailFolder subfolder in (IEnumerable<IMailFolder>) folder1.GetSubfolders(false, cancellationToken))
        {
          if (subfolder.Name == "[Gmail]" || subfolder.Name == "PROCESSED" || subfolder.Name == "NoXmlFound")
          {
            Console.WriteLine("[folder] {0} Omitido", (object) subfolder.Name);
          }
          else
          {
            Console.WriteLine("[folder] {0}", (object) subfolder.Name);
            IMailFolder folder2 = imapClient.GetFolder(subfolder.Name, new CancellationToken());
            try
            {
              int num3 = (int) folder2.Open(FolderAccess.ReadWrite, new CancellationToken());
            }
            catch (Exception ex)
            {
              Console.WriteLine(ex.ToString());
              continue;
            }
            foreach (UniqueId uniqueId in (IEnumerable<UniqueId>) folder2.Search((SearchQuery) binarySearchQuery, new CancellationToken()))
            {
              bool bXml = false;
              MimeMessage message = folder2.GetMessage(uniqueId, new CancellationToken(), (ITransferProgress) null);
              Console.WriteLine("[match] {0}: {1}", (object) uniqueId, (object) message.Subject);
              IEnumerable<MimeEntity> mimeEntities = message.Attachments;
              if (mimeEntities == null || mimeEntities.Count<MimeEntity>() == 0)
                mimeEntities = (IEnumerable<MimeEntity>) message.BodyParts.Where<MimeEntity>((Func<MimeEntity, bool>) (x => x.ContentType != null && x is MimePart)).ToList<MimeEntity>();
              IEnumerable<MimeEntity> second = mimeEntities.OfType<TnefPart>().SelectMany<TnefPart, MimeEntity>((Func<TnefPart, IEnumerable<MimeEntity>>) (a => a.ExtractAttachments()));
              foreach (MimeEntity mimeEntity in mimeEntities.Concat<MimeEntity>(second))
              {
                if (mimeEntity is MessagePart)
                {
                  Console.WriteLine("[attachment] {0}", (object) ((object) (MessagePart) mimeEntity).ToString());
                }
                else
                {
                  MimePart mimePart = (MimePart) mimeEntity;
                  Console.WriteLine("[attachment] {0} {1}", (object) mimePart.ContentType.MimeType, (object) mimePart.FileName);
                  if (mimePart.ContentType.MimeType == "text/xml" || mimePart.ContentType.MimeType == "application/xml" || mimePart.ContentType.MimeType == "application/octet-stream" && mimePart.FileName.ToUpper().EndsWith(".XML"))
                  {
                    try
                    {
                      MemoryStream stream = new MemoryStream();
                      mimePart.Content.DecodeTo((Stream) stream, new CancellationToken());
                      CfdiHeaderVO cfdiHeaderVo = Program.ReadXml(stream, ref bXml);
                      if (cfdiHeaderVo != null)
                      {
                        if (cfdiHeaderVo.UUID != null)
                        {
                          if (cfdiHeaderVo.UUID != "")
                            cfdiHeaderVoList.Add(cfdiHeaderVo);
                        }
                      }
                    }
                    catch (Exception ex)
                    {
                      Console.Write(ex.ToString());
                    }
                  }
                  else if (mimePart.ContentType.MimeType == "application/zip" || mimePart.ContentType.MimeType == "application/octet-stream" && mimePart.FileName.ToUpper().EndsWith(".ZIP"))
                  {
                    try
                    {
                      MemoryStream memoryStream = new MemoryStream();
                      mimePart.Content.DecodeTo((Stream) memoryStream, new CancellationToken());
                      memoryStream.Position = 0L;
                      using (ZipFile zipFile = ZipFile.Read((Stream) memoryStream))
                      {
                        foreach (ZipEntry zipEntry in zipFile)
                        {
                          Console.WriteLine("Content file: {0} {1}", (object) zipEntry.FileName, (object) ((object) zipEntry).GetType());
                          if (zipEntry.FileName.ToUpper().EndsWith(".XML"))
                          {
                            MemoryStream stream = new MemoryStream();
                            zipEntry.Extract((Stream) stream);
                            CfdiHeaderVO cfdiHeaderVo = Program.ReadXml(stream, ref bXml);
                            if (cfdiHeaderVo != null && cfdiHeaderVo.UUID != null && cfdiHeaderVo.UUID != "")
                              cfdiHeaderVoList.Add(cfdiHeaderVo);
                          }
                        }
                      }
                    }
                    catch (Exception ex)
                    {
                      Console.Write(ex.ToString());
                    }
                  }
                }
              }
              if (bXml)
              {
                if (!stringList1.Contains(folder2.Name))
                  folder2.SetLabels(uniqueId, (IList<string>) stringList1, false, new CancellationToken());
                folder2.RemoveLabels(uniqueId, (IList<string>) stringList2, false, new CancellationToken());
              }
              else if (!stringList2.Contains(folder2.Name))
                folder2.SetLabels(uniqueId, (IList<string>) stringList2, false, new CancellationToken());
            }
          }
        }
      }
      if (cfdiHeaderVoList.Count <= 0)
        return;
      using (StreamWriter streamWriter = new StreamWriter(path1))
      {
        foreach (CfdiHeaderVO cfdiHeaderVo in cfdiHeaderVoList)
          streamWriter.WriteLine(cfdiHeaderVo.ToString());
      }
      using (StreamWriter streamWriter = new StreamWriter(path2))
      {
        foreach (CfdiHeaderVO cfdiHeaderVo in cfdiHeaderVoList)
        {
          foreach (CfdiRelatedVO cfdiRelatedVo in cfdiHeaderVo.lsRelated)
            streamWriter.WriteLine(cfdiHeaderVo.UUID + "\t" + cfdiRelatedVo.trelat + "\t" + cfdiRelatedVo.uuid + "\t" + cfdiRelatedVo.fechaPago + "\t" + cfdiRelatedVo.montoPago);
        }
      }
    }

    private static CfdiHeaderVO ReadXml(MemoryStream stream, ref bool bXml)
    {
      CfdiHeaderVO cfdiHeaderVo = (CfdiHeaderVO) null;
      string xml = Encoding.UTF8.GetString(stream.ToArray());
      string str1 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
      if (xml.StartsWith(str1))
        xml = xml.Remove(0, str1.Length);
      XmlDocument xmlDocument = new XmlDocument();
      xmlDocument.LoadXml(xml);
      XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDocument.NameTable);
      nsmgr.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
      nsmgr.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");
      nsmgr.AddNamespace("pago10", "http://www.sat.gob.mx/Pagos");
      foreach (XmlNode selectNode1 in xmlDocument.SelectNodes("//cfdi:Comprobante", nsmgr))
      {
        cfdiHeaderVo = new CfdiHeaderVO();
        string nodeAttribute1 = Program.getNodeAttribute(selectNode1, "TipoDeComprobante");
        string nodeAttribute2 = Program.getNodeAttribute(selectNode1, "SubTotal");
        string nodeAttribute3 = Program.getNodeAttribute(selectNode1, "Total");
        string nodeAttribute4 = Program.getNodeAttribute(selectNode1, "Descuento");
        string nodeAttribute5 = Program.getNodeAttribute(selectNode1, "Version");
        string nodeAttribute6 = Program.getNodeAttribute(selectNode1, "Serie");
        string nodeAttribute7 = Program.getNodeAttribute(selectNode1, "Folio");
        string nodeAttribute8 = Program.getNodeAttribute(selectNode1, "Fecha");
        Console.WriteLine("     Tipo: {0}, Subtotal: {1}, Total: {2}", (object) nodeAttribute1, (object) nodeAttribute2, (object) nodeAttribute3);
        string str2 = "";
        string str3 = "";
        string str4 = "";
        string str5 = "";
        foreach (XmlNode selectNode2 in selectNode1.SelectNodes("cfdi:Complemento/tfd:TimbreFiscalDigital", nsmgr))
        {
          bXml = true;
          str2 = Program.getNodeAttribute(selectNode2, "UUID");
          str3 = Program.getNodeAttribute(selectNode2, "FechaTimbrado");
          str4 = Program.getNodeAttribute(selectNode2, "NoCertificadoSAT");
          str5 = Program.getNodeAttribute(selectNode2, "SelloSAT");
          Console.WriteLine("     UUID: {0}", (object) str2);
        }
        string str6 = "";
        foreach (XmlNode selectNode3 in selectNode1.SelectNodes("cfdi:Emisor", nsmgr))
        {
          str6 = Program.getNodeAttribute(selectNode3, "Rfc");
          Console.WriteLine("     RfcEmisor: {0}", (object) str6);
        }
        string str7 = "";
        foreach (XmlNode selectNode4 in selectNode1.SelectNodes("cfdi:Receptor", nsmgr))
        {
          str7 = Program.getNodeAttribute(selectNode4, "Rfc");
          Console.WriteLine("     RfcReceptor: {0}", (object) str7);
        }
        cfdiHeaderVo.RfcR = str7;
        cfdiHeaderVo.RfcE = str6;
        cfdiHeaderVo.UUID = str2;
        cfdiHeaderVo.Type = nodeAttribute1;
        cfdiHeaderVo.version = nodeAttribute5;
        cfdiHeaderVo.doc_prefix = nodeAttribute6;
        cfdiHeaderVo.doc_number = nodeAttribute7;
        cfdiHeaderVo.inv_date_time = nodeAttribute8;
        cfdiHeaderVo.cfdi_date_time = str3;
        cfdiHeaderVo.cert_sat = str4;
        cfdiHeaderVo.seal_sat = str5;
        cfdiHeaderVo.subtotal = nodeAttribute2;
        cfdiHeaderVo.total = nodeAttribute3;
        cfdiHeaderVo.discount = nodeAttribute4;
        foreach (XmlNode selectNode5 in selectNode1.SelectNodes("cfdi:CfdiRelacionados", nsmgr))
        {
          string nodeAttribute9 = Program.getNodeAttribute(selectNode5, "TipoRelacion");
          Console.WriteLine("     Tipo Relacion: {0}", (object) nodeAttribute9);
          selectNode1.SelectNodes("cfdi:CfdiRelacionado", nsmgr);
          foreach (XmlNode node in selectNode5)
          {
            string nodeAttribute10 = Program.getNodeAttribute(node, "UUID");
            Console.WriteLine("     UUID Relacionado: {0}", (object) nodeAttribute10);
            cfdiHeaderVo.lsRelated.Add(new CfdiRelatedVO()
            {
              trelat = nodeAttribute9,
              uuid = nodeAttribute10
            });
          }
        }
        if (nodeAttribute1 == "P")
        {
          foreach (XmlNode selectNode6 in selectNode1.SelectNodes("cfdi:Complemento/pago10:Pagos", nsmgr))
          {
            foreach (XmlNode selectNode7 in selectNode6.SelectNodes("pago10:Pago", nsmgr))
            {
              string nodeAttribute11 = Program.getNodeAttribute(selectNode7, "FechaPago");
              foreach (XmlNode selectNode8 in selectNode7.SelectNodes("pago10:DoctoRelacionado", nsmgr))
              {
                string nodeAttribute12 = Program.getNodeAttribute(selectNode8, "IdDocumento");
                string nodeAttribute13 = Program.getNodeAttribute(selectNode8, "ImpPagado");
                Console.WriteLine("     IdDocumento Relacionado: {0}", (object) nodeAttribute12);
                cfdiHeaderVo.lsRelated.Add(new CfdiRelatedVO()
                {
                  trelat = "",
                  uuid = nodeAttribute12,
                  fechaPago = nodeAttribute11,
                  montoPago = nodeAttribute13
                });
              }
            }
          }
        }
      }
      return cfdiHeaderVo;
    }

    private static string getNodeAttribute(XmlNode node, string attr)
    {
      string nodeAttribute = "";
      XmlAttribute attribute = node.Attributes[attr];
      if (attribute != null)
        nodeAttribute = attribute.Value;
      return nodeAttribute;
    }
  }
}
