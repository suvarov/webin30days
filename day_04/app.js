function getData() {
  var conn = new ActiveXObject("ADODB.Connection") ;

  var connStr = "Data Source=KZALAWS003;Initial Catalog=CityDelivery;User ID=telebot;Password=5jnUX4b5CtqWjHGm;Provider=SQLOLEDB";

  conn.Open(connStr);
  var rs = new ActiveXObject("ADODB.Recordset");

  rs.Open("SELECT * FROM CityDelivery.dbo.vw_ADUserInfo", conn);
  rs.MoveFirst
  while(!rs.eof)
  {
     document.write('<span>' + rs.fields(1) + '</span><br>');
     rs.movenext;
  }

  rs.close;
  conn.close;
}
