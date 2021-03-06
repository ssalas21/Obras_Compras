//------------------------------------------------------------------------------
// <auto-generated>
//    Este código se generó a partir de una plantilla.
//
//    Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//    Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Data.EntityClient;
using System.Data.Objects;
using System.Data.Objects.DataClasses;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml.Serialization;

[assembly: EdmSchemaAttribute()]
namespace Permisos_de_obras
{
    #region Contextos
    
    /// <summary>
    /// No hay documentación de metadatos disponible.
    /// </summary>
    public partial class BaseEntities : ObjectContext
    {
        #region Constructores
    
        /// <summary>
        /// Inicializa un nuevo objeto BaseEntities usando la cadena de conexión encontrada en la sección 'BaseEntities' del archivo de configuración de la aplicación.
        /// </summary>
        public BaseEntities() : base("name=BaseEntities", "BaseEntities")
        {
            this.ContextOptions.LazyLoadingEnabled = true;
            OnContextCreated();
        }
    
        /// <summary>
        /// Inicializar un nuevo objeto BaseEntities.
        /// </summary>
        public BaseEntities(string connectionString) : base(connectionString, "BaseEntities")
        {
            this.ContextOptions.LazyLoadingEnabled = true;
            OnContextCreated();
        }
    
        /// <summary>
        /// Inicializar un nuevo objeto BaseEntities.
        /// </summary>
        public BaseEntities(EntityConnection connection) : base(connection, "BaseEntities")
        {
            this.ContextOptions.LazyLoadingEnabled = true;
            OnContextCreated();
        }
    
        #endregion
    
        #region Métodos parciales
    
        partial void OnContextCreated();
    
        #endregion
    
        #region Propiedades de ObjectSet
    
        /// <summary>
        /// No hay documentación de metadatos disponible.
        /// </summary>
        public ObjectSet<Utm> Utm
        {
            get
            {
                if ((_Utm == null))
                {
                    _Utm = base.CreateObjectSet<Utm>("Utm");
                }
                return _Utm;
            }
        }
        private ObjectSet<Utm> _Utm;

        #endregion

        #region Métodos AddTo
    
        /// <summary>
        /// Método desusado para agregar un nuevo objeto al EntitySet Utm. Considere la posibilidad de usar el método .Add de la propiedad ObjectSet&lt;T&gt; asociada.
        /// </summary>
        public void AddToUtm(Utm utm)
        {
            base.AddObject("Utm", utm);
        }

        #endregion

    }

    #endregion

    #region Entidades
    
    /// <summary>
    /// No hay documentación de metadatos disponible.
    /// </summary>
    [EdmEntityTypeAttribute(NamespaceName="BaseModel", Name="Utm")]
    [Serializable()]
    [DataContractAttribute(IsReference=true)]
    public partial class Utm : EntityObject
    {
        #region Método de generador
    
        /// <summary>
        /// Crear un nuevo objeto Utm.
        /// </summary>
        /// <param name="idUtm">Valor inicial de la propiedad IdUtm.</param>
        public static Utm CreateUtm(global::System.Int32 idUtm)
        {
            Utm utm = new Utm();
            utm.IdUtm = idUtm;
            return utm;
        }

        #endregion

        #region Propiedades simples
    
        /// <summary>
        /// No hay documentación de metadatos disponible.
        /// </summary>
        [EdmScalarPropertyAttribute(EntityKeyProperty=true, IsNullable=false)]
        [DataMemberAttribute()]
        public global::System.Int32 IdUtm
        {
            get
            {
                return _IdUtm;
            }
            set
            {
                if (_IdUtm != value)
                {
                    OnIdUtmChanging(value);
                    ReportPropertyChanging("IdUtm");
                    _IdUtm = StructuralObject.SetValidValue(value, "IdUtm");
                    ReportPropertyChanged("IdUtm");
                    OnIdUtmChanged();
                }
            }
        }
        private global::System.Int32 _IdUtm;
        partial void OnIdUtmChanging(global::System.Int32 value);
        partial void OnIdUtmChanged();
    
        /// <summary>
        /// No hay documentación de metadatos disponible.
        /// </summary>
        [EdmScalarPropertyAttribute(EntityKeyProperty=false, IsNullable=true)]
        [DataMemberAttribute()]
        public Nullable<global::System.Int32> Anno
        {
            get
            {
                return _Anno;
            }
            set
            {
                OnAnnoChanging(value);
                ReportPropertyChanging("Anno");
                _Anno = StructuralObject.SetValidValue(value, "Anno");
                ReportPropertyChanged("Anno");
                OnAnnoChanged();
            }
        }
        private Nullable<global::System.Int32> _Anno;
        partial void OnAnnoChanging(Nullable<global::System.Int32> value);
        partial void OnAnnoChanged();
    
        /// <summary>
        /// No hay documentación de metadatos disponible.
        /// </summary>
        [EdmScalarPropertyAttribute(EntityKeyProperty=false, IsNullable=true)]
        [DataMemberAttribute()]
        public Nullable<global::System.Int32> Mes
        {
            get
            {
                return _Mes;
            }
            set
            {
                OnMesChanging(value);
                ReportPropertyChanging("Mes");
                _Mes = StructuralObject.SetValidValue(value, "Mes");
                ReportPropertyChanged("Mes");
                OnMesChanged();
            }
        }
        private Nullable<global::System.Int32> _Mes;
        partial void OnMesChanging(Nullable<global::System.Int32> value);
        partial void OnMesChanged();
    
        /// <summary>
        /// No hay documentación de metadatos disponible.
        /// </summary>
        [EdmScalarPropertyAttribute(EntityKeyProperty=false, IsNullable=true)]
        [DataMemberAttribute()]
        public Nullable<global::System.Int32> Valor
        {
            get
            {
                return _Valor;
            }
            set
            {
                OnValorChanging(value);
                ReportPropertyChanging("Valor");
                _Valor = StructuralObject.SetValidValue(value, "Valor");
                ReportPropertyChanged("Valor");
                OnValorChanged();
            }
        }
        private Nullable<global::System.Int32> _Valor;
        partial void OnValorChanging(Nullable<global::System.Int32> value);
        partial void OnValorChanged();

        #endregion

    }

    #endregion

}
