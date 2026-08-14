# Plataforma institucional COINVIERTE

Primera versión en Streamlit + Supabase para gestionar proyectos institucionales.

## Ejecutar localmente

1. Instala dependencias: `pip install -r requirements.txt`
2. Copia `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y agrega las credenciales de Supabase.
3. Ejecuta `supabase_schema.sql` en el SQL Editor de Supabase.
4. En Supabase Authentication habilita Email/Password y crea usuarios `@jalisco.gob.mx`.
5. Inicia: `streamlit run app.py`

Sin credenciales de Supabase, la app funciona en modo demostración y permite revisar el flujo visual.

## Logo

Crea la carpeta `assets` y coloca el logo como `assets/logo_coinvierte.png`.

## Seguridad

La restricción de dominio se valida tanto en la interfaz como mediante políticas RLS en Supabase. El bucket `expedientes` es privado.

## Conexión y gestión de usuarios

1. Crea o abre el proyecto de Supabase destinado a esta plataforma.
2. En **SQL Editor**, ejecuta todo el archivo `supabase_schema.sql`.
3. En **Authentication > Providers > Email**, habilita correo y contraseña. Para este piloto se recomienda desactivar temporalmente la confirmación obligatoria de correo; el código entregado por el administrador funciona como validación de acceso.
4. En **Authentication > Users**, crea la cuenta administradora `yani.limberopulos@jalisco.gob.mx` y establece una contraseña segura. El esquema vinculará esta cuenta automáticamente con el rol administrador.
5. En **Project Settings > API Keys**, copia la clave pública `sb_publishable_...` (o la clave `anon` en proyectos antiguos). Nunca copies la clave secreta o `service_role` en GitHub.
6. En Streamlit Cloud abre **Manage app > Settings > Secrets** y registra:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_TU-CLAVE-PUBLICA"
```

7. Reinicia la aplicación e ingresa con la cuenta administradora. En el menú aparecerá **Gestión de usuarios**.

### Flujo de alta

- El administrador captura nombre y correo institucional y genera un código temporal.
- El código vence, se guarda cifrado y sólo puede utilizarse una vez.
- La persona abre **Activar acceso con código**, captura su correo, código y nueva contraseña.
- El administrador puede suspender, reactivar o emitir un nuevo código.
