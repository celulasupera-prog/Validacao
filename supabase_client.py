from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from urllib import parse, request


class SupabaseError(RuntimeError):
    pass


@dataclass
class SupabaseRestClient:
    url: str
    api_key: str

    def _request(
        self,
        method: str,
        path: str,
        params: Optional[Dict[str, Any]] = None,
        body: Optional[Any] = None,
        prefer: Optional[str] = None,
    ) -> Any:
        query = ""
        if params:
            query = "?" + parse.urlencode(params, doseq=True)

        endpoint = f"{self.url.rstrip('/')}/rest/v1/{path}{query}"
        headers = {
            "apikey": self.api_key,
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        if prefer:
            headers["Prefer"] = prefer

        payload = None if body is None else json.dumps(body).encode("utf-8")
        req = request.Request(endpoint, data=payload, headers=headers, method=method.upper())

        try:
            with request.urlopen(req, timeout=20) as response:
                data = response.read().decode("utf-8")
                if not data:
                    return None
                return json.loads(data)
        except Exception as exc:  # pragma: no cover
            raise SupabaseError(f"Erro ao acessar Supabase em {path}: {exc}") from exc

    def ensure_default_groups(self, names: List[str]) -> None:
        existentes = self._request("GET", "grupos", {"select": "nome"})
        nomes_existentes = {r.get("nome") for r in existentes or []}
        faltantes = [nome for nome in names if nome not in nomes_existentes]
        if faltantes:
            payload = [{"nome": nome, "ativo": True} for nome in faltantes]
            self._request("POST", "grupos", body=payload, prefer="return=representation")

    def get_groups(self) -> List[Dict[str, Any]]:
        grupos = self._request("GET", "grupos", {"select": "id,nome,ativo", "order": "nome.asc"})
        return grupos or []

    def get_group_records(self, table: str, group_id: int) -> List[Dict[str, Any]]:
        registros = self._request(
            "GET",
            table,
            {
                "select": "id,codigo_empresa,nome_empresa,codigo_empregado,nome_empregado,ativo",
                "grupo_id": f"eq.{group_id}",
                "order": "codigo_empresa.asc,codigo_empregado.asc",
            },
        )
        return registros or []

    def sync_group_records(self, table: str, group_id: int, rows: List[Dict[str, Any]]) -> None:
        atuais = self.get_group_records(table, group_id)
        ids_atuais = {int(r["id"]) for r in atuais if r.get("id") is not None}

        ids_novos = set()
        inserts: List[Dict[str, Any]] = []

        for row in rows:
            codigo_empresa = str(row.get("codigo_empresa", "")).strip()
            nome_empresa = str(row.get("nome_empresa", "")).strip()
            codigo_empregado = str(row.get("codigo_empregado", "")).strip()
            nome_empregado = str(row.get("nome_empregado", "")).strip()
            ativo = bool(row.get("ativo", True))

            if not (codigo_empresa and nome_empresa and codigo_empregado and nome_empregado):
                continue

            payload = {
                "grupo_id": group_id,
                "codigo_empresa": codigo_empresa,
                "nome_empresa": nome_empresa,
                "codigo_empregado": codigo_empregado,
                "nome_empregado": nome_empregado,
                "ativo": ativo,
            }

            row_id = row.get("id")
            if row_id in (None, "", 0, "0"):
                inserts.append(payload)
                continue

            row_id = int(row_id)
            ids_novos.add(row_id)
            self._request(
                "PATCH",
                table,
                params={"id": f"eq.{row_id}", "grupo_id": f"eq.{group_id}"},
                body=payload,
                prefer="return=minimal",
            )

        if inserts:
            self._request("POST", table, body=inserts, prefer="return=minimal")

        ids_para_deletar = ids_atuais - ids_novos
        if ids_para_deletar:
            ids_str = ",".join(str(i) for i in sorted(ids_para_deletar))
            self._request(
                "DELETE",
                table,
                params={"id": f"in.({ids_str})", "grupo_id": f"eq.{group_id}"},
                prefer="return=minimal",
            )
