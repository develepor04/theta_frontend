# PMO - Project Management Office

A modern React-based frontend application for processing EDDR (Engineering Document Delivery Record) files with intelligent automation and beautiful visualizations.

## 🚀 Features

### Core Functionality
- **Smart File Processing**: Upload EDDR Excel files and get instant processing with intelligent data extraction
- **Custom Output Columns**: Select and customize output columns, add custom fields on the fly
- **History & Analytics**: Track all processed files with detailed history
- **Subscription Management**: Free, Pro, and Enterprise tiers with different limits

### User Experience
- **Modern UI**: Clean, professional interface inspired by Smartsheet
- **Smooth Animations**: GSAP-powered animations for delightful interactions
- **Premium Blue Theme**: White and premium blue color scheme
- **Responsive Design**: Works seamlessly on desktop, tablet, and mobile

### Pages
1. **Landing Page**: Marketing page with features and pricing
2. **Authentication**: Login and Sign-up pages
3. **Dashboard**: Main workspace with sidebar, file uploader, and column selector
4. **History**: View all past processing jobs
5. **Output**: Detailed view of processed data with tables
6. **Subscription**: Manage subscription plans and view usage

## 🛠️ Tech Stack

- **React 18** - UI library
- **Vite** - Build tool and dev server
- **React Router** - Client-side routing
- **Zustand** - State management
- **GSAP** - Animation library
- **Lucide React** - Icon library
- **Axios** - HTTP client

## 📦 Installation

### Prerequisites
- Node.js (v16 or higher)
- npm or yarn

### Setup Steps

1. **Navigate to the frontend directory**:
   ```bash
   cd frontend
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Start the development server**:
   ```bash
   npm run dev
   ```

4. **Open your browser**:
   Navigate to `http://localhost:3000`

## 📁 Project Structure

```
frontend/
├── public/                 # Static assets
├── src/
│   ├── pages/             # Page components
│   │   ├── LandingPage.jsx
│   │   ├── LoginPage.jsx
│   │   ├── SignUpPage.jsx
│   │   ├── Dashboard.jsx
│   │   ├── History.jsx
│   │   ├── Output.jsx
│   │   └── Subscription.jsx
│   ├── store/             # State management
│   │   └── useStore.js
│   ├── App.jsx            # Main app component
│   ├── main.jsx           # Entry point
│   └── index.css          # Global styles
├── index.html
├── package.json
└── vite.config.js
```

## 🎨 Design System

### Colors
- **Primary Blue**: `#0073ea`
- **Dark Blue**: `#0060b9`
- **Light Blue**: `#e6f3ff`
- **White**: `#ffffff`
- **Text Dark**: `#323338`
- **Text Gray**: `#676879`

### Components
- Buttons: Primary and Secondary variants
- Cards: Elevated containers with hover effects
- Form inputs: Clean with icon support
- Tables: Responsive with hover states

## 🔐 Authentication

Current implementation uses mock authentication for demonstration. In production:

1. Replace mock login in `LoginPage.jsx` and `SignUpPage.jsx` with API calls
2. Implement JWT token management
3. Add secure token storage
4. Integrate with backend authentication service

## 💰 Subscription Model

### Free Plan
- 1 file upload per day
- Up to 3 total uploads
- Basic output columns
- 7-day history

### Pro Plan ($49/month)
- Unlimited uploads
- Custom columns
- Advanced analytics
- Unlimited history
- Priority support

### Enterprise Plan (Custom pricing)
- Everything in Pro
- SSO & SAML
- API access
- Dedicated support
- Custom integrations

## 🚀 Building for Production

```bash
npm run build
```

The built files will be in the `dist` directory.

## 🔗 Backend Integration

To connect with the Python backend (`App.py`):

1. **Create API endpoints** in a Flask/FastAPI application
2. **Update the proxy** in `vite.config.js` to point to your backend URL
3. **Implement file upload** endpoint that accepts Excel files
4. **Process files** using the existing `extract_eddr_data` function
5. **Return JSON** with processed data

### Example API Structure

```python
# Flask example
@app.route('/api/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    # Process using extract_eddr_data()
    return jsonify({'status': 'success', 'data': processed_data})
```

## 📝 Environment Variables

Create a `.env` file in the frontend directory:

```env
VITE_API_URL=http://localhost:5000
VITE_APP_NAME=PMO
```

## 🧪 Testing

```bash
npm run test
```

## 📄 License

This project is proprietary software.

## 👥 Support

For support, contact your development team or open an issue in the project repository.

## 🎯 Next Steps

1. **Backend Integration**: Connect to Python backend API
2. **Authentication**: Implement real authentication with JWT
3. **Payment Gateway**: Integrate Stripe/PayPal for subscriptions
4. **File Processing**: Connect upload to actual EDDR processing
5. **Data Visualization**: Add charts and graphs for analytics
6. **Export Functionality**: Implement Excel/PDF export
7. **Testing**: Add unit and integration tests
8. **Deployment**: Deploy to production (Vercel, Netlify, etc.)

---

Built with ❤️ for efficient EDDR data management

# Theta — Development and Deployment guide

## 1. Development Setup Document
[Development Document](https://apliaglobal77-my.sharepoint.com/:w:/g/personal/aniket_apliaglobal_com/IQDD-RS2lxbQT7aByRl5JKtWAVoyNIKxbygDjxty-iomcNI?e=d5VT5F)

## 2. Deployment Document
[Deployment Document](https://apliaglobal77-my.sharepoint.com/:w:/g/personal/aniket_apliaglobal_com/IQD3oXE_AP86TrlNOTCX_hyXAQ7CKGhzC-H-g7Ql-ER1qp4?e=ZUiTRS)

## 3. Files for Data Ingetion:
[Data Files](https://apliaglobal77-my.sharepoint.com/:f:/g/personal/aniket_apliaglobal_com/IgDwtfrxXJizQJwCTEnnEmFUAaLIxde285860Boh82h-mDA?e=RCkpdQ)

