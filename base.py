import json
from typing import List, Optional
from sqlalchemy import create_engine, String, Integer, Float, ForeignKey
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, Session

# 1. Definición de Modelos ORM (Sintaxis SQLAlchemy 2.0)
class Base(DeclarativeBase):
    pass

class Marca(Base):
    __tablename__ = "marca"
    
    id_marca: Mapped[int] = mapped_column(primary_key=True)
    nombre: Mapped[str] = mapped_column(String(50), unique=True)
    origen: Mapped[Optional[str]] = mapped_column(String(50))
    
    # Relación uno a muchos (carga en cascada)
    modelos: Mapped[List["Modelo"]] = relationship(back_populates="marca", cascade="all, delete-orphan")

class Modelo(Base):
    __tablename__ = "modelo"
    
    id_modelo: Mapped[int] = mapped_column(primary_key=True)
    id_marca: Mapped[int] = mapped_column(ForeignKey("marca.id_marca"))
    nombre: Mapped[str] = mapped_column(String(100))
    tipo_carroceria: Mapped[Optional[str]] = mapped_column(String(30))
    
    marca: Mapped["Marca"] = relationship(back_populates="modelos")
    especificaciones: Mapped[List["EspecificacionTecnica"]] = relationship(back_populates="modelo", cascade="all, delete-orphan")

class EspecificacionTecnica(Base):
    __tablename__ = "especificacion_tecnica"
    
    id_especificacion: Mapped[int] = mapped_column(primary_key=True)
    id_modelo: Mapped[int] = mapped_column(ForeignKey("modelo.id_modelo"))
    version: Mapped[Optional[str]] = mapped_column(String(100))
    autonomia_km: Mapped[Optional[float]] = mapped_column(Float)
    bateria_kwh: Mapped[Optional[float]] = mapped_column(Float)
    potencia_hp: Mapped[Optional[int]] = mapped_column(Integer)
    aceleracion_0_100: Mapped[Optional[float]] = mapped_column(Float)
    
    modelo: Mapped["Modelo"] = relationship(back_populates="especificaciones")

# 2. Configuración del Motor (Engine)
# Nota: Para Oracle cambia a 'oracle+cx_oracle://usuario:pass@host:puerto/?service_name=nombre'
engine = create_engine("sqlite:///vehiculos.db", echo=False) 
Base.metadata.create_all(engine)

# 3. Lógica de Ingesta
def poblar_base_datos(ruta_archivo: str):
    with open(ruta_archivo, 'r', encoding='utf-8') as f:
        datos_json = json.load(f)
        
    with Session(engine) as session:
        for marca_data in datos_json:
            # 3.1 Construir Entidad Padre
            nueva_marca = Marca(
                nombre=marca_data["marca"],
                origen=marca_data.get("origen")
            )
            
            # 3.2 Construir y anidar Entidades Hijas (Modelos)
            for modelo_data in marca_data.get("modelos", []):
                nuevo_modelo = Modelo(
                    nombre=modelo_data["nombre"],
                    tipo_carroceria=modelo_data.get("tipo_carroceria")
                )
                nueva_marca.modelos.append(nuevo_modelo)
                
                # 3.3 Construir y anidar Entidades Nietas (Especificaciones)
                for spec_data in modelo_data.get("versiones", []):
                    nueva_spec = EspecificacionTecnica(
                        version=spec_data.get("version"),
                        autonomia_km=spec_data.get("autonomia_km"),
                        bateria_kwh=spec_data.get("bateria_kwh"),
                        potencia_hp=spec_data.get("potencia_hp"),
                        aceleracion_0_100=spec_data.get("aceleracion_0_100")
                    )
                    nuevo_modelo.especificaciones.append(nueva_spec)
            
            # Solo necesitas agregar el padre; SQLAlchemy infiere y propaga las FK al hacer commit
            session.add(nueva_marca)
            
        session.commit()
        print("Datos ingeridos y relaciones construidas con éxito.")

# 4. Validar la carga ejecutando una consulta (Query)
def verificar_carga():
    from sqlalchemy import select
    with Session(engine) as session:
        # Ejecutar un JOIN entre las tres tablas
        stmt = select(Marca.nombre, Modelo.nombre, EspecificacionTecnica.autonomia_km)\
            .join(Marca.modelos)\
            .join(Modelo.especificaciones)\
            .where(EspecificacionTecnica.autonomia_km >= 400)
            
        resultados = session.execute(stmt).all()
        
        print("\nVehículos con autonomía >= 400 km:")
        for marca, modelo, autonomia in resultados:
            print(f"- {marca} {modelo}: {autonomia} km")

# Ejecución
if __name__ == "__main__":
    poblar_base_datos('ecuador.json')
    verificar_carga()