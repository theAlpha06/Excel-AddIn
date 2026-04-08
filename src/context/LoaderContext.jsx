import { AnimationStyles } from "@fluentui/react";
import { TvUsbRegular } from "@fluentui/react-icons";
import React, { createContext, useState } from "react";

export const LoaderContext = createContext();

export const LoaderProvider = ({ children }) => {
    const [isLoading, setIsLoading] = useState(false);

    const value = {
        isLoading,
        setIsLoading
    };

    return (
        <LoaderContext.Provider value={value}>
            {children}
        </LoaderContext.Provider>
    );
};