using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Koanvi.Data.DataTable {
  public class DataTable : System.Data.DataTable {
    public DataTable():base() { }
  }
  public class DataSet : System.Data.DataSet {
    public DataSet() : base() { }
  }
}

namespace Koanvi.Data.DataAdapter {
  public class SqlDataAdapter : System.ComponentModel.Component {
    public System.Data.SqlClient.SqlDataAdapter Adapter;

    public SqlDataAdapter() : base() {
      Adapter = new System.Data.SqlClient.SqlDataAdapter();
    }

    
  }
  public class IDbDataAdapter : System.Data.IDbDataAdapter {
    public IDbDataAdapter() : base() { }

    public IDbCommand DeleteCommand {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public IDbCommand InsertCommand {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public MissingMappingAction MissingMappingAction {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public MissingSchemaAction MissingSchemaAction {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public IDbCommand SelectCommand {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public ITableMappingCollection TableMappings {
      get {
        throw new NotImplementedException();
      }
    }

    public IDbCommand UpdateCommand {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public Int32 Fill(DataSet dataSet) {
      throw new NotImplementedException();
    }

    public System.Data.DataTable[] FillSchema(DataSet dataSet, SchemaType schemaType) {
      throw new NotImplementedException();
    }

    public IDataParameter[] GetFillParameters() {
      throw new NotImplementedException();
    }

    public Int32 Update(DataSet dataSet) {
      throw new NotImplementedException();
    }
  }
  public class IDataAdapter : System.Data.IDataAdapter {
    public MissingMappingAction MissingMappingAction {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public MissingSchemaAction MissingSchemaAction {
      get {
        throw new NotImplementedException();
      }

      set {
        throw new NotImplementedException();
      }
    }

    public ITableMappingCollection TableMappings {
      get {
        throw new NotImplementedException();
      }
    }

    public Int32 Fill(DataSet dataSet) {
      throw new NotImplementedException();
    }

    public System.Data.DataTable[] FillSchema(DataSet dataSet, SchemaType schemaType) {
      throw new NotImplementedException();
    }

    public IDataParameter[] GetFillParameters() {
      throw new NotImplementedException();
    }

    public Int32 Update(DataSet dataSet) {
      throw new NotImplementedException();
    }
  }

}