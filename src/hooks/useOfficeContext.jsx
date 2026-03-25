import React from 'react';
import { OfficeContextProvider } from '../context/OfficeContext';

export const useOfficeContext = () => {
    const context = React.useContext(OfficeContextProvider);
    if (!context) {
        throw new Error('useOfficeContext must be used within OfficeProvider');
    }
    return context;
};