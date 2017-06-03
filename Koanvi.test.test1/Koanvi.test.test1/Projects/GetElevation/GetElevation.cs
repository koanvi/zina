using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Koanvi.Projects.GetElevation {

  public class GetElevation {


    //пример использования:

    //var GetElevation = new Koanvi.Projects.GetElevation.GetElevation();

    //GetElevation.GetResult(55.824409, 37.447417);


    //  GetElevation.GetResultSquare(
    //    new Projects.GetElevation.Location() { lat = 55.824409, lng = 37.447417 },
    //    new Projects.GetElevation.Location() { lat = 58.824409, lng = 39.447417 }
    //    );

    public GetElevation() {}

    public Elevation GetResult(Location location) {
      return GetResult(location.lat, location.lng);
    }
    public Elevation GetResult(double lat, double lng) {

      var retval = GetResult(CreateHttpWebRequest(lat, lng));
      return retval;

    }
    public Elevation GetResult(System.Net.HttpWebRequest HttpWebRequest) {

      System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(Elevation));

      try {
        return (Elevation)serializer.ReadObject(HttpWebRequest.GetResponse().GetResponseStream());
      } catch(System.Net.WebException e) {
        return (Elevation)serializer.ReadObject(e.Response.GetResponseStream());
      }
    }

    public System.Net.HttpWebRequest CreateHttpWebRequest(double lat, double lng) {

      //http://0s.nvqxa4y.m5xw6z3mmvqxa2ltfzrw63i.cmle.ru/maps/api/elevation/json?locations=61.54,59.70

      string RequestText = $@"http://0s.nvqxa4y.m5xw6z3mmvqxa2ltfzrw63i.cmle.ru/maps/api/elevation/json?locations="
      + lat.ToString(System.Globalization.CultureInfo.InvariantCulture)
      +","
      + lng.ToString(System.Globalization.CultureInfo.InvariantCulture);

      return (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(RequestText);

    }

    public List<Elevation> GetResultSquare(Location StartLocation, Location EndLocation) {

      // делаем список
      // бежим по списку
      // добавляем 2 координаты для каждого ээлемента (справа и снизу)


      List<LocationElevation> result = new List<LocationElevation>();

      result.Add(new LocationElevation() {location= StartLocation });

      result.ForEach(curResult => {

        curResult.elevation=GetResult(curResult.location);
        result.Add(new LocationElevation() {
          location =new Location() { lat= curResult.location.lat
          , lng= curResult.location.lng+ Location.km_2_lng(curResult.elevation.results[0].resolution/1000, curResult.location.lat)
          }
        });
      });

      int cur = 0;
      //while(CurLocation.lat < EndLocation.lat && CurLocation.lng < EndLocation.lng) {}

      return result.Select(x=>x.elevation).ToList();

    }

  }

  public class Elevation {
    public Result[] results { get; set; }
    public string status { get; set; }
  }
  public class Result {
    public string elevation { get; set; }
    public Location location { get; set; }
    public double resolution { get; set; }
  }
  public class Location {
    //Широта  - Lat(Y)	= (с севера	на юг) -90 до +90 постоянно 1 градус = 111 км.
    //Долгота - lng(X)	= (идет с запада на восток)-180 до +180 1 градус = lng_2_km(Lat)
    public double lat { get; set; }
    public double lng { get; set; }

    #region static

    public static int EARTH_RADIUS = 6371;

    /// <summary>
    /// в одном lng км 
    /// </summary>
    /// <param name="lat"></param>
    /// <returns></returns>
    public static double one_lng_2_km(double lat) {
      //EARTH_RADIUS * (Math.PI / 180) - 111
      return EARTH_RADIUS * (Math.PI / 180) * Math.Cos(lat * Math.PI / 180);
    }

    /// <summary>
    /// в одном lat км 
    /// </summary>
    /// <returns></returns>
    public static double one_lat_2_km() {
      return EARTH_RADIUS * (Math.PI / 180);
    }

    public static double km_2_lng(double km, double lat) {
      return one_lng_2_km(lat) * km;
    }

    public static double km_2_lat(double km) {
      return one_lat_2_km()* km;
    }

    #endregion

  }
  public class LocationElevation {
    public Location location { get;set;}
    public Elevation elevation { get; set; }

  }
}
