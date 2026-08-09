import React, { useState } from 'react';
import {
  CheckCircle,
  Zap,
  FileSpreadsheet,
  Briefcase,
  Building2,
  Menu,
  X,
} from 'lucide-react';
import Sidebar from '../components/Sidebar';
import './Subscription.css';

const ACTIVE_PLAN_ID = 'enterprise';

const plans = [
  {
    id: 'free',
    name: 'Free Access',
    price: 'Meeting signup',
    period: '5-day access',
    icon: Zap,
    description: 'Try a simple sheet after booking a meeting.',
    features: [
      'Simple sheet: 20 rows × 5 columns',
      'Must sign up for a meeting',
      '5-day access window',
    ],
  },
  {
    id: 'entry',
    name: 'Entry Level',
    price: '1,200',
    period: '/annum',
    icon: FileSpreadsheet,
    description: 'Uploaded Google Sheets only — no Theta Sheets.',
    features: [
      'Access to uploaded Google Sheets',
      'No Theta Sheets',
      'Professional Services (hours package)',
      'License rate: 1,200 per annum',
    ],
  },
  {
    id: 'professional',
    name: 'Professional',
    price: '1,500',
    period: '/annum',
    icon: Briefcase,
    description: 'Professional licenses with Theta Sheets access.',
    features: [
      'Professional licenses',
      'Access to Theta Sheets',
      'Professional Services available',
      '1,500 per annum',
    ],
  },
  {
    id: 'enterprise',
    name: 'Enterprise',
    price: 'Custom',
    period: '',
    icon: Building2,
    description: 'Full Professional capability plus system integrations.',
    features: [
      'Everything in Professional',
      'Integration to Primavera (or other systems)',
      'Professional Services available',
    ],
  },
];

const Subscription = () => {
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  return (
    <div className="subscription-page">
      <button
        className="mobile-menu-button"
        onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
        aria-label="Toggle menu"
      >
        {isMobileMenuOpen ? <X size={24} /> : <Menu size={24} />}
      </button>

      <div
        className={`mobile-sidebar-overlay ${isMobileMenuOpen ? 'open' : ''}`}
        onClick={() => setIsMobileMenuOpen(false)}
      />

      <Sidebar isMobileMenuOpen={isMobileMenuOpen} setIsMobileMenuOpen={setIsMobileMenuOpen} />

      <div className="main-content">
        <div className="subscription-container">
          <div className="subscription-header">
            <div>
              <h1>Theta Pulse Light</h1>
              <p>Your current active plan and available tiers</p>
            </div>
          </div>

          <div className="pricing-section">
            <div className="section-header">
              <h2>Plans</h2>
              <p>Enterprise is your active license</p>
            </div>

            <div className="pricing-grid pulse-light-grid">
              {plans.map((plan) => {
                const isActive = plan.id === ACTIVE_PLAN_ID;
                const Icon = plan.icon;

                return (
                  <div
                    key={plan.id}
                    className={`pricing-card card ${isActive ? 'current' : ''}`}
                  >
                    {isActive && <div className="badge-current">Active</div>}

                    <div className="pricing-header">
                      <div
                        className="pricing-icon"
                        style={{
                          background: isActive
                            ? 'rgba(16, 185, 129, 0.12)'
                            : 'rgba(148, 163, 184, 0.12)',
                          color: isActive ? '#059669' : '#64748b',
                        }}
                      >
                        <Icon size={26} />
                      </div>
                      <h3>{plan.name}</h3>
                      <p className="plan-card-desc">{plan.description}</p>
                      <div className="pricing-price">
                        <span className="amount">{plan.price}</span>
                        {plan.period && <span className="period">{plan.period}</span>}
                      </div>
                    </div>

                    <ul className="pricing-features">
                      {plan.features.map((feature) => (
                        <li key={feature}>
                          <CheckCircle size={16} />
                          <span>{feature}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Subscription;
