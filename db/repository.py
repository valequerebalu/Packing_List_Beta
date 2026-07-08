"""Consultas y operaciones de persistencia para la base local."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path
from typing import Any

from sqlalchemy.orm import Session

from db.models import Customer, Generation, GenerationStatus, SharePointSource, Template, User


def get_or_create_user(session: Session, *, email: str, full_name: str, microsoft_id: str | None = None, tenant_id: str | None = None) -> User:
	user = session.query(User).filter(User.email == email).one_or_none()
	if user is not None:
		user.full_name = full_name or user.full_name
		user.microsoft_id = microsoft_id or user.microsoft_id
		user.tenant_id = tenant_id or user.tenant_id
		return user

	user = User(email=email, full_name=full_name, microsoft_id=microsoft_id, tenant_id=tenant_id)
	session.add(user)
	session.flush()
	return user


def get_or_create_customer(session: Session, *, code: str, name: str, country: str | None = None) -> Customer:
	customer = session.query(Customer).filter(Customer.code == code).one_or_none()
	if customer is not None:
		customer.name = name or customer.name
		customer.country = country or customer.country
		return customer

	customer = Customer(code=code, name=name, country=country)
	session.add(customer)
	session.flush()
	return customer


def get_or_create_template(
	session: Session,
	*,
	customer_id: int,
	created_by: int,
	name: str,
	file_path: str,
	version: str = "1.0.0",
	mapping_config: str | None = None,
) -> Template:
	template = (
		session.query(Template)
		.filter(Template.customer_id == customer_id, Template.name == name, Template.version == version)
		.one_or_none()
	)
	if template is not None:
		template.file_path = file_path or template.file_path
		template.mapping_config = mapping_config or template.mapping_config
		return template

	template = Template(
		customer_id=customer_id,
		created_by=created_by,
		name=name,
		file_path=file_path,
		version=version,
		mapping_config=mapping_config,
	)
	session.add(template)
	session.flush()
	return template


def sha256_file(file_path: str | Path) -> str:
	digest = sha256()
	with Path(file_path).open("rb") as handle:
		for chunk in iter(lambda: handle.read(1024 * 1024), b""):
			digest.update(chunk)
	return digest.hexdigest()


def record_generation(
	session: Session,
	*,
	user_id: int,
	customer_id: int,
	template_id: int,
	packing_file_name: str,
	thermograph_file_name: str,
	packing_file_hash: str | None = None,
	thermograph_file_hash: str | None = None,
	sharepoint_source_id: int | None = None,
	output_file_name: str | None = None,
	output_file_path: str | None = None,
	status: GenerationStatus = GenerationStatus.PENDING,
	error_detail: str | None = None,
	duration_seconds: int | None = None,
) -> Generation:
	generation = Generation(
		user_id=user_id,
		customer_id=customer_id,
		template_id=template_id,
		sharepoint_source_id=sharepoint_source_id,
		packing_file_name=packing_file_name,
		packing_file_hash=packing_file_hash,
		thermograph_file_name=thermograph_file_name,
		thermograph_file_hash=thermograph_file_hash,
		output_file_name=output_file_name,
		output_file_path=output_file_path,
		status=status,
		error_detail=error_detail,
		duration_seconds=duration_seconds,
	)
	session.add(generation)
	session.flush()
	return generation


def save_generation_success(
	session: Session,
	generation: Generation,
	*,
	output_file_name: str,
	output_file_path: str,
	duration_seconds: int | None = None,
) -> Generation:
	generation.status = GenerationStatus.SUCCESS
	generation.output_file_name = output_file_name
	generation.output_file_path = output_file_path
	generation.duration_seconds = duration_seconds
	session.add(generation)
	session.flush()
	return generation


def save_generation_error(session: Session, generation: Generation, error_detail: str, *, duration_seconds: int | None = None) -> Generation:
	generation.status = GenerationStatus.ERROR
	generation.error_detail = error_detail
	generation.duration_seconds = duration_seconds
	session.add(generation)
	session.flush()
	return generation


def get_or_create_sharepoint_source(
	session: Session,
	*,
	file_name: str,
	file_url: str,
	file_version: str | None = None,
	last_modified=None,
) -> SharePointSource:
	source = session.query(SharePointSource).filter(SharePointSource.file_name == file_name, SharePointSource.file_url == file_url).one_or_none()
	if source is not None:
		source.file_version = file_version or source.file_version
		source.last_modified = last_modified or source.last_modified
		return source

	source = SharePointSource(file_name=file_name, file_url=file_url, file_version=file_version, last_modified=last_modified)
	session.add(source)
	session.flush()
	return source
