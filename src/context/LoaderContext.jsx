import React, { createContext, useState } from "react";

export const LoaderContext = createContext();

export const LoaderProvider = ({ children }) => {
    const [isLoading, setIsLoading] = useState(false);

    const value = {
        isLoading,
        setIsLoading,
        show: () => setIsLoading(true),
        hide: () => setIsLoading(false),
    };

    return (
        <LoaderContext.Provider value={value}>
            {children}
        </LoaderContext.Provider>
    );
};
