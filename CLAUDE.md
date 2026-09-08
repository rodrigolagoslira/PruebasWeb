# CLAUDE.md - Instrucciones de Proyecto

Este archivo contiene las instrucciones y principios que Claude debe seguir al inicio de cada proyecto de desarrollo.

---

## Preferencias del Usuario (rlagos@scraices.cl)

### Directorio de salida de documentos generados
Todo archivo Excel, PDF o documento generado debe guardarse en:
```
C:\Users\rodri\OneDrive\Documentos Claude Code
```
En sesiones remotas (donde esa ruta no existe), usar `output/` local como fallback y avisar al usuario que el archivo quedó en el servidor remoto.

---

## Principios de Desarrollo

### 1. Arquitectura y Diseño

**Clean Architecture**
- Separar claramente las capas: UI, Logica de Negocio, Datos
- Los componentes de UI no deben contener logica de negocio
- Usar inversión de dependencias para desacoplar módulos
- Preferir composición sobre herencia

**SOLID Principles**
- **S**ingle Responsibility: Cada módulo/clase tiene una sola razón para cambiar
- **O**pen/Closed: Abierto para extensión, cerrado para modificación
- **L**iskov Substitution: Las clases derivadas deben ser sustituibles por sus bases
- **I**nterface Segregation: Interfaces específicas mejor que una general
- **D**ependency Inversion: Depender de abstracciones, no de implementaciones

**DRY (Don't Repeat Yourself)**
- Extraer código duplicado a funciones/componentes reutilizables
- Centralizar constantes y configuraciones
- Usar hooks personalizados para lógica repetida

### 2. Estructura del Código

**Organización de Carpetas**
```
src/
├── app/              # Rutas y páginas (Next.js App Router)
├── components/       # Componentes React
│   ├── ui/          # Componentes base reutilizables
│   └── [feature]/   # Componentes específicos por funcionalidad
├── hooks/           # Custom hooks
├── lib/             # Utilidades y clientes externos
├── types/           # Definiciones TypeScript
├── constants/       # Constantes y configuraciones
└── store/           # Estado global (Zustand/Context)
```

**Convenciones de Nombres**
- Componentes: PascalCase (ej: `UserProfile.tsx`)
- Hooks: camelCase con prefijo "use" (ej: `useAuth.ts`)
- Utilidades: camelCase (ej: `formatDate.ts`)
- Constantes: UPPER_SNAKE_CASE (ej: `MAX_ITEMS`)
- Tipos/Interfaces: PascalCase (ej: `interface User {}`)

### 3. Calidad del Código

**TypeScript**
- Usar tipos estrictos, evitar `any`
- Definir interfaces para props de componentes
- Exportar tipos para reutilización
- Usar generics cuando sea apropiado

**Manejo de Errores**
- Implementar error boundaries para componentes
- Usar try/catch para operaciones async
- Mostrar mensajes de error amigables al usuario
- Loggear errores para debugging

**Performance**
- Usar `useMemo` y `useCallback` para optimización
- Implementar lazy loading para componentes pesados
- Virtualizar listas largas
- Optimizar imágenes y assets

### 4. Testing

**Estrategia de Testing**
- Unit tests para funciones y hooks
- Integration tests para flujos de usuario
- E2E tests para caminos críticos

**Herramientas Recomendadas**
- Vitest para unit tests
- Testing Library para componentes React
- Playwright para E2E
- MSW para mocking de APIs

**Cobertura Mínima**
- Funciones utilitarias: 100%
- Hooks personalizados: 80%
- Componentes críticos: 70%

### 5. Git y Versionado

**Commits**
- Mensajes descriptivos y concisos
- Prefijos: feat, fix, docs, style, refactor, test, chore
- Un commit = un cambio lógico

**Branches**
- main: código en producción
- develop: integración de features
- feature/[nombre]: nuevas funcionalidades
- fix/[nombre]: correcciones

### 6. Documentación

**Código**
- Documentar funciones complejas con JSDoc
- Comentarios solo cuando el código no es autoexplicativo
- README actualizado con instrucciones de setup

**Proyecto**
- Documentacion.md con explicación técnica
- Changelog para versiones
- Diagramas para arquitectura compleja

---

## Flujo de Trabajo con Claude

### Al Iniciar un Proyecto

1. **Leer especificaciones** - Explorar documentos de requisitos en la carpeta
2. **Crear Todo List** - Usar TodoWrite para planificar tareas
3. **Diseñar arquitectura** - Definir estructura antes de codificar
4. **Setup inicial** - Configurar proyecto base con dependencias

### Durante el Desarrollo

1. **Actualizar Todo List** - Marcar tareas completadas inmediatamente
2. **Commits incrementales** - No acumular muchos cambios
3. **Probar mientras se desarrolla** - No dejar testing para el final
4. **Documentar decisiones** - Anotar el "por qué" de decisiones importantes

### Al Finalizar

1. **Verificación completa** - Probar todos los flujos principales
2. **Limpieza** - Remover código comentado y console.logs
3. **Documentacion.md** - Escribir explicación técnica del proyecto
4. **Revisión de Todo List** - Confirmar que todo está completado

---

## Template de Documentacion.md

```markdown
# Documentación Técnica: [Nombre del Proyecto]

## Resumen
Breve descripción del proyecto y su propósito.

## Stack Tecnológico
- Frontend: [tecnologías]
- Backend: [tecnologías]
- Base de datos: [tecnología]
- Otros: [herramientas]

## Arquitectura
Descripción de la estructura del proyecto y decisiones arquitectónicas.

## Módulos Principales
### Módulo 1
- Propósito
- Archivos principales
- Dependencias

### Módulo 2
...

## Flujos de Usuario
Descripción de los principales flujos de la aplicación.

## Configuración
Variables de entorno y configuraciones necesarias.

## Comandos
- `npm run dev` - Desarrollo
- `npm run build` - Build de producción
- `npm run test` - Ejecutar tests

## Notas Adicionales
Consideraciones especiales, limitaciones conocidas, trabajo pendiente.
```

---

## Checklist de Calidad

Antes de marcar un proyecto como completado:

- [ ] Código compila sin errores
- [ ] No hay warnings de TypeScript
- [ ] Tests pasan correctamente
- [ ] Responsive funciona en móvil y desktop
- [ ] Accesibilidad básica implementada
- [ ] Variables de entorno documentadas
- [ ] README actualizado
- [ ] Documentacion.md creado
- [ ] Todo List completado
