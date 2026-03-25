import React, {createContext, useState, useEffect} from 'react';

export const AuthContext = createContext(null);

export const AuthProvider = ({children}) => {
    const [user, setUser] = useState(null);
    const [isConnected, setIsConnected] = useState(false);

    const login = (userData) => {
        setUser(userData);
    }

    const logout = () => {
        setUser(null);
    }

    const value = {
        user,
        setUser,
        login,
        logout,
        isConnected,
        setIsConnected
    }

    return (
        <AuthContext.Provider value={value}>
            {children}
        </AuthContext.Provider>
    );
};