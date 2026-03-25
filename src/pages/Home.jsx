import React, { useState } from 'react';
import Profile from '../components/Profile';

const Home = () => {
    const [activeTab, setActiveTab] = useState('profile');

    const tabButtonStyle = (isActive) => ({
        padding: '16px 28px',
        border: 'none',
        backgroundColor: 'transparent',
        color: isActive ? '#0078d4' : '#666',
        cursor: 'pointer',
        fontSize: '15px',
        fontWeight: '500',
        transition: 'all 0.3s ease',
        position: 'relative',
        outline: 'none',
        overflow: 'hidden'
    });

    const tabUnderlineStyle = (isActive) => ({
        position: 'absolute',
        bottom: 0,
        left: 0,
        right: 0,
        height: '3px',
        backgroundColor: '#0078d4',
        transform: isActive ? 'scaleX(1)' : 'scaleX(0)',
        transformOrigin: isActive ? 'left' : 'right',
        transition: 'transform 0.4s cubic-bezier(0.4, 0, 0.2, 1)'
    });

    return (
        <div style={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column', backgroundColor: '#f9f9f9' }}>
            <div style={{ 
                display: 'flex', 
                borderBottom: '1px solid #e5e5e5',
                backgroundColor: '#fff',
                boxShadow: '0 1px 3px rgba(0, 0, 0, 0.05)'
            }}>
                <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                    <button
                        onClick={() => setActiveTab('profile')}
                        style={tabButtonStyle(activeTab === 'profile')}
                    >
                        Profile
                        <div style={tabUnderlineStyle(activeTab === 'profile')} />
                    </button>
                </div>
                <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                    <button
                        onClick={() => setActiveTab('salesforce')}
                        style={tabButtonStyle(activeTab === 'salesforce')}
                    >
                        Salesforce Sync
                        <div style={tabUnderlineStyle(activeTab === 'salesforce')} />
                    </button>
                </div>
            </div>

            <div style={{ 
                flex: 1, 
                overflow: 'auto',
                animation: 'fadeIn 0.5s ease-in-out',
                padding: '10px'
            }}>
                {activeTab === 'profile' && <Profile />}
                {activeTab === 'salesforce' && <div style={{ padding: '20px' }}>Salesforce Sync Content</div>}
            </div>
        </div>
    );
}

export default Home;