using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Koanvi.Projects.HtmlHoover {
  /// <summary>
  /// 2 варианта работы:
  /// берем данные из входного списка
  /// берем данные из страницы
  /// </summary>
  public class HtmlHoover {
    public List<Model.RequestResult> requestResult { get;set;}
    public HtmlHoover(List<string> Requests) {
      requestResult = new List<Model.RequestResult>();
      Requests.ForEach(x=> {
        requestResult.Add(new Model.RequestResult(x));
      });
    }//public HtmlHoover(List<string> Requests)
    public void Fill() {
      requestResult.ForEach(x => {
        x.Fill();
      });
    }//public void Fill()
  }//public class HtmlHoover
}//namespace Koanvi.Projects.HtmlHoover
namespace Koanvi.Projects.HtmlHoover.Model {

  public class RequestResult {
    private Koanvi.Net.HttpWebRequest _HttpWebRequest;
    public Koanvi.Net.HttpWebRequest HttpWebRequest { get { return _HttpWebRequest; }  }
    private string _ResponseString;
    public string ResponseString { get { return _ResponseString; } }
    /// <summary>
    /// Номер страницы
    /// </summary>
    public int Page { get; set; }
    /// <summary>
    /// что именно надо брать
    /// </summary>
    public ParseResult parseResult { get; }
    public RequestResult(String requestUriString) {
      this._HttpWebRequest = new Net.HttpWebRequest(requestUriString);
    }
    public void Fill() {
      _ResponseString=this._HttpWebRequest.GetResponseString();
      
    }
  }
  public class ParseResult {
    public String Text { get; set; }
    public ParseResult(String Text) {
      this.Text = Text;
    }
  }
}
namespace Koanvi.Net {

  public class HttpWebRequest {
    public System.Net.HttpWebRequest Request;
    public HttpWebRequest(String requestUriString) {
      Request = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(requestUriString);
    }
    public System.IO.Stream GetResponseStream() {
      try {
        return Request.GetResponse().GetResponseStream();
      } catch(System.Net.WebException e) {
        return e.Response.GetResponseStream();
      }
    }
    public string GetResponseString() {
      string responseText;
      using(var reader = new System.IO.StreamReader(GetResponseStream())) {
        responseText = reader.ReadToEnd();
      }
      return responseText;
    }
  }//public class HttpWebRequest


}