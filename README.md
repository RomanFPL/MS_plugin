# Outlook Reply Generator Add-in

React-based Microsoft Outlook Add-in that generates email replies using external API.

## 🚀 Features

- **API Integration**: Send email subject and body to external API for reply generation
- **React + TypeScript**: Modern development stack with full type safety
- **Tailwind CSS**: Beautiful, responsive UI design
- **Office.js**: Full integration with Outlook APIs
- **Development Mode**: Test outside Outlook environment

## 📋 Prerequisites

- Node.js 16+ 
- npm or yarn
- Microsoft Outlook (Desktop or Web)
- Office Add-ins development environment

## 🛠️ Installation

1. **Clone the repository**
```bash
git clone <your-repo-url>
cd MS_plagin
```

2. **Install dependencies**
```bash
npm install
```

3. **Start development server**
```bash
npm run dev
```

Server will start at `http://localhost:3000`

## 🔧 Development

### Project Structure
```
├── src/
│   ├── taskpane/
│   │   ├── TaskPane.tsx     # Main React component
│   │   ├── TaskPane.css     # Tailwind styles
│   │   └── index.tsx        # Entry point
│   ├── vite-env.d.ts        # TypeScript types
│   └── types/
├── manifest.xml             # Development manifest (localhost)
├── manifest.prod.xml        # Production manifest (HTTPS)
├── vite.config.ts           # Vite configuration
└── tailwind.config.js       # Tailwind configuration
```

### Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run start` - Start dev server with browser

### Testing the Add-in

**Option 1: Browser Testing**
1. Open `http://localhost:3000`
2. Test with mock data (works without Outlook)

**Option 2: Outlook Integration**
1. Start development server: `npm run dev`
2. In Outlook, go to Developer settings
3. Load `manifest.xml` file
4. Add-in will appear in Outlook interface

## 📤 Production Deployment

1. **Build the project**
```bash
npm run build
```

2. **Upload `dist/` folder to your web server**

3. **Update manifest URLs**
   - Use `manifest.prod.xml`
   - Replace URLs with your production domain
   - Ensure HTTPS is used

4. **Validate manifest**
```bash
npx office-addin-validator manifest.prod.xml
```

## 🏗️ Tech Stack

- **Frontend**: React 18, TypeScript
- **Styling**: Tailwind CSS
- **Build Tool**: Vite
- **Office Integration**: Office.js
- **Development**: Hot reload, TypeScript checking

## ⚙️ Configuration

### API Integration
The add-in sends POST requests to your API with this structure:
```typescript
{
  subject: string,      // Email subject
  bodyPreview: string   // First 5000 chars of email body
}
```

### Environment Variables
Create `.env.local` for custom configuration:
```env
VITE_API_URL=https://your-api.com/generate
```

## 🐛 Troubleshooting

### Common Issues

1. **"This site can't provide a secure connection"**
   - Accept the self-signed certificate in browser
   - Or use HTTP in development (already configured)

2. **Office.js not loading**
   - Check if running in Outlook context
   - Use browser testing for development

3. **Manifest validation errors**
   - Use `manifest.prod.xml` for production
   - Ensure all URLs use HTTPS
   - Check icon dimensions (64x64, 128x128)

### Validation
```bash
# Validate development manifest
npx office-addin-validator manifest.xml

# Validate production manifest  
npx office-addin-validator manifest.prod.xml
```

## 📝 License

MIT License - see LICENSE file for details
