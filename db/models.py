from sqlalchemy import (
    Column,
    Integer,
    String,
    Boolean,
    DateTime,
    Text,
    ForeignKey,
    Enum as SAEnum,
)
from sqlalchemy.orm import relationship
from datetime import datetime
import enum
from db.database import Base

# ─── Enums ────────────────────────────────────────────────────────────────────


class GenerationStatus(enum.Enum):
    SUCCESS = "success"
    ERROR = "error"
    PENDING = "pending"


class TemplateStatus(enum.Enum):
    ACTIVE = "active"
    INACTIVE = "inactive"
    DRAFT = "draft"


# ─── Models ───────────────────────────────────────────────────────────────────


class User(Base):
    """
    Representa a un usuario de la empresa Beta que tiene acceso al sistema de generación de Packing List.
    """

    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    email = Column(String(255), unique=True, nullable=True, index=True)
    full_name = Column(String(255), nullable=False)
    is_active = Column(Boolean, default=True, nullable=False)
    is_admin = Column(Boolean, default=False, nullable=False)  # can manage templates
    last_login = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    generations = relationship("Generation", back_populates="user", lazy="dynamic")
    templates = relationship(
        "Template", back_populates="created_by_user", lazy="dynamic"
    )

    def __repr__(self):
        return f"<User {self.username}>"


class Customer(Base):
    """
    Clientes exportadores de Beta. Cada cliente puede tener una o más plantillas.
    Se desacopla de Template para que un cliente pueda tener múltiples versiones
    de plantilla sin necesidad de duplicar información.
    """

    __tablename__ = "customers"

    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(50), unique=True, nullable=False, index=True)  # e.g. "CLI_001"
    name = Column(String(255), nullable=False)
    country = Column(String(100), nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Relationships
    templates = relationship("Template", back_populates="customer", lazy="dynamic")
    generations = relationship("Generation", back_populates="customer", lazy="dynamic")

    def __repr__(self):
        return f"<Customer {self.code} - {self.name}>"


class Template(Base):
    """
    Plantillas de Excel por cliente. Soporta versionado y múltiples plantillas por cliente.
    Almacena la configuración de mapeo de campos (JSON) para escalar flexiblemente a 20+ plantillas.
    """

    __tablename__ = "templates"

    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(
        Integer, ForeignKey("customers.id"), nullable=False, index=True
    )
    created_by = Column(Integer, ForeignKey("users.id"), nullable=False)
    name = Column(String(255), nullable=False)
    version = Column(String(20), default="1.0.0", nullable=False)  # semver
    file_path = Column(String(500), nullable=False)  # path to .xlsx
    mapping_config = Column(Text, nullable=True)  # JSON field mapping
    status = Column(SAEnum(TemplateStatus), default=TemplateStatus.ACTIVE)
    notes = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Relationships
    customer = relationship("Customer", back_populates="templates")
    created_by_user = relationship("User", back_populates="templates")
    generations = relationship("Generation", back_populates="template", lazy="dynamic")

    def __repr__(self):
        return f"<Template {self.name} v{self.version} - {self.customer}>"


class SharePointSource(Base):
    """
    Registra las versiones de archivos de SharePoint utilizados durante
    cada generación. Permite auditar exactamente qué datos de la nube
    se usaron para producir cada Excel de salida.
    """

    __tablename__ = "sharepoint_sources"

    id = Column(Integer, primary_key=True, index=True)
    file_name = Column(String(255), nullable=False)
    file_url = Column(String(500), nullable=False)
    file_version = Column(String(100), nullable=True)  # SP version tag
    last_modified = Column(DateTime, nullable=True)  # from SP metadata
    fetched_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    # Relationships
    generations = relationship("Generation", back_populates="sharepoint_source")

    def __repr__(self):
        return f"<SharePointSource {self.file_name} @ {self.file_version}>"


class Generation(Base):
    """
    Registro completo de auditoría de cada evento de generación de Excel.
    Enlaza usuario, cliente, plantilla, archivos de entrada y fuente de SharePoint.
    Diseñado para rastreabilidad y reportes futuros.
    """

    __tablename__ = "generations"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False, index=True)
    customer_id = Column(
        Integer, ForeignKey("customers.id"), nullable=False, index=True
    )
    template_id = Column(Integer, ForeignKey("templates.id"), nullable=False)
    sharepoint_source_id = Column(
        Integer, ForeignKey("sharepoint_sources.id"), nullable=True
    )

    # Input files (store name + hash for integrity)
    packing_file_name = Column(String(255), nullable=False)
    packing_file_hash = Column(String(64), nullable=True)  # SHA-256
    thermograph_file_name = Column(String(255), nullable=False)
    thermograph_file_hash = Column(String(64), nullable=True)  # SHA-256

    # Output
    output_file_name = Column(String(255), nullable=True)
    output_file_path = Column(String(500), nullable=True)

    # Status & traceability
    status = Column(SAEnum(GenerationStatus), default=GenerationStatus.PENDING)
    error_detail = Column(Text, nullable=True)
    duration_seconds = Column(Integer, nullable=True)  # processing time
    generated_at = Column(DateTime, default=datetime.utcnow, nullable=False, index=True)

    # Relationships
    user = relationship("User", back_populates="generations")
    customer = relationship("Customer", back_populates="generations")
    template = relationship("Template", back_populates="generations")
    sharepoint_source = relationship("SharePointSource", back_populates="generations")

    def __repr__(self):
        return f"<Generation #{self.id} | {self.customer} | {self.status}>"
