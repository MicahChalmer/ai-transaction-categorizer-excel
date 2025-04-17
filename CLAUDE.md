# AI Autocategorizer for Excel - Development Guidelines

## Build & Development Commands
- `npm run build` - Production build
- `npm run build:dev` - Development build
- `npm run dev-server` - Start webpack development server
- `npm run start` - Start the add-in in Excel
- `npm run watch` - Start incremental build in watch mode
- `npm run lint` - Check for linting issues
- `npm run lint:fix` - Fix linting issues
- `npm run prettier` - Run prettier code formatter

## Code Style Guidelines
- **TypeScript**: Use TypeScript for all new files with strict type checking
- **React Components**: Use functional components with hooks, explicit interface props
- **Imports**: Group imports (React, components, interfaces, icons/utils)
- **Naming**: PascalCase for components & interfaces, camelCase for variables & functions
- **CSS**: Use makeStyles from Fluent UI for styling with tokens
- **Error Handling**: Use async/await with try/catch blocks for error handling
- **Component Structure**: Props interface -> styles -> component -> exports
- **State Management**: Prefer useState and useContext for state
- **Formatting**: Follows office-addin-prettier-config