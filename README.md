# Robot Panadería — Proyecto Final de Maestría

## Descripción del proyecto

Este repositorio contiene el desarrollo tecnológico de un proyecto académico de emprendimiento social panadero, propuesto para el distrito de Bebedero de Cañas, Guanacaste, Costa Rica.

El proyecto forma parte del trabajo final del **Máster en Marketing Digital y Big Data de la Universidad Internacional Iberoamericana (UNINI)**.

Su objetivo es diseñar una estrategia integral de marketing digital y una arquitectura de datos que permitan mejorar la toma de decisiones, la planificación de la producción, el control financiero y la sostenibilidad de un pequeño emprendimiento rural con propósito social.

## Arquitectura tecnológica

El sistema integra diferentes componentes:

- **Robot Python:** automatiza la captura y el procesamiento de información externa mediante APIs.
- **Excel Maestro:** modelo financiero y operativo que integra inventarios, producción, ventas, costos y proyecciones.
- **Aplicación web SaaS:** interfaz para gestionar operaciones, pedidos, inventarios y registros del negocio.
- **WhatsApp Business y Twilio:** integración para la gestión de pedidos digitales.
- **Google Drive:** almacenamiento e intercambio de archivos de trabajo.
- **GitHub Actions:** ejecución programada de procesos automatizados.
- **Render:** alojamiento de la aplicación web.

La arquitectura está organizada en cuatro capas: ingesta, procesamiento, almacenamiento y comunicación de datos.

## Funciones principales

El modelo tecnológico contempla:

- Captura de indicadores externos, como tipo de cambio, clima y feriados.
- Gestión de materias primas y productos elaborados.
- Registro de hornadas y producción.
- Control de ventas presenciales y pedidos digitales.
- Gestión de inventarios y costos.
- Integración de información operativa con el modelo financiero.
- Generación de información para apoyar decisiones de producción y comercialización.

## Estado del proyecto

El robot Python, la aplicación SaaS y los componentes de integración digital han sido desarrollados y sometidos a pruebas funcionales dentro del alcance académico.

La puesta en marcha comercial de la panadería es una etapa futura, sujeta a financiamiento, permisos y demás condiciones de implementación.

Las campañas de Meta Ads y los tableros de Power BI forman parte de la estrategia de implementación prevista.

## Propósito social

El emprendimiento propone reinvertir una parte de sus excedentes en formación técnica para los habitantes de Bebedero, mediante una futura vinculación con el Instituto Nacional de Aprendizaje (INA).

La finalidad es combinar sostenibilidad financiera, innovación tecnológica y desarrollo local.

## Seguridad y configuración

Las credenciales y configuraciones sensibles utilizadas por los servicios externos deben administrarse mediante variables de entorno y mecanismos de almacenamiento protegido, como GitHub Actions Secrets.

Los archivos de ejemplo de configuración no deben contener credenciales reales.

Este repositorio se presenta con fines de documentación y evaluación académica.

## Autoría

**Francisco Javier Fernández Ulloa**

Proyecto Final de Máster en Marketing Digital y Big Data.

Universidad Internacional Iberoamericana (UNINI).

**Costa Rica, 2026.**
