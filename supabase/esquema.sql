-- Esquema para Supabase: historial de solicitudes de ingreso.
-- Ejecutar en Supabase → SQL Editor → New query → pegar y ejecutar.

create table if not exists public.solicitudes (
    id                 bigint primary key,
    caso               text,
    oficina            text,
    solicitante        text,
    fecha_solicitud    text,
    regional           text,
    aprobador          text,
    estado             text not null default 'PENDIENTE APROBACIÓN',
    fecha_proceso      text,
    fecha_respuesta    text,
    actualizado_por    text,
    actualizado_en     timestamptz not null default now()
);

create index if not exists solicitudes_estado_idx on public.solicitudes (estado);
create index if not exists solicitudes_regional_idx on public.solicitudes (regional);
create index if not exists solicitudes_actualizado_idx on public.solicitudes (actualizado_en desc);

create table if not exists public.historial_estados (
    id             bigserial primary key,
    solicitud_id   bigint references public.solicitudes (id) on delete cascade,
    estado_anterior text,
    estado_nuevo    text not null,
    usuario         text,
    observacion     text,
    creado_en       timestamptz not null default now()
);

create index if not exists historial_solicitud_idx on public.historial_estados (solicitud_id);

-- Seguridad a nivel de fila.
-- La consulta pública lee; las escrituras solo pasan con la clave service_role,
-- que omite RLS y nunca se expone en el navegador.
alter table public.solicitudes enable row level security;
alter table public.historial_estados enable row level security;

drop policy if exists "lectura publica solicitudes" on public.solicitudes;
create policy "lectura publica solicitudes"
    on public.solicitudes for select
    using (true);

drop policy if exists "lectura auditoria" on public.historial_estados;
create policy "lectura auditoria"
    on public.historial_estados for select
    using (true);
