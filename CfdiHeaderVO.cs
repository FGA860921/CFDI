// Decompiled with JetBrains decompiler
// Type: ReadCFDIfromMail.CfdiHeaderVO
// Assembly: ReadCFDIfromMail, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 5C3D3298-27B9-4099-8001-CBF5C018B0E0
// Assembly location: C:\tmp\ReadCFDIfromMail\ReadCFDIfromMail.dll

using System.Collections.Generic;

namespace ReadCFDIfromMail
{
  internal class CfdiHeaderVO
  {
    public List<CfdiRelatedVO> lsRelated = new List<CfdiRelatedVO>();

    public string RfcR { get; set; }

    public string RfcE { get; set; }

    public string UUID { get; set; }

    public string Type { get; set; }

    public string version { get; set; }

    public string doc_prefix { get; set; }

    public string doc_number { get; set; }

    public string inv_date_time { get; set; }

    public string cfdi_date_time { get; set; }

    public string cert_sat { get; set; }

    public string seal_sat { get; set; }

    public string subtotal { get; set; }

    public string total { get; set; }

    public string discount { get; set; }

    public override string ToString() => this.RfcR + "\t" + this.RfcE + "\t" + this.UUID + "\t" + this.Type + "\t" + this.version + "\t" + this.doc_prefix + "\t" + this.doc_number + "\t" + this.inv_date_time + "\t" + this.cfdi_date_time + "\t" + this.subtotal + "\t" + this.total + "\t" + this.discount + "\t" + this.cert_sat + "\t" + this.seal_sat;
  }
}
