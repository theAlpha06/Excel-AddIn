import React, { createContext, useState, useEffect } from 'react';

export const OfficeContextProvider = createContext();

export const OfficeProvider = ({ children }) => {
    const [isOfficeReady, setIsOfficeReady] = useState(false);
    const [officeError, setOfficeError] = useState(null);

    useEffect(() => {
        const initializeOffice = async () => {
            return new Promise((resolve) => {
                if (typeof Office === 'undefined') {
                    setOfficeError('Office library not loaded');
                    resolve(false);
                    return;
                }

                Office.onReady((info) => {
                    try {
                        if (info && info.host && Office.context) {
                            setIsOfficeReady(true);
                            setOfficeError(null);
                            resolve(true);
                        } else {
                            setOfficeError('Office context not initialized');
                            resolve(false);
                        }
                    } catch (error) {
                        setOfficeError(error.message);
                        resolve(false);
                    }
                });

                setTimeout(() => {
                    if (!isOfficeReady) {
                        try {
                            if (typeof Office !== 'undefined' && Office.context) {
                                setIsOfficeReady(true);
                                resolve(true);
                            }
                        } catch (error) {
                            setOfficeError('Office context initialization timeout');
                            resolve(false);
                        }
                    }
                }, 2000);
            });
        };

        initializeOffice();
    }, []);

    return (
        <OfficeContextProvider.Provider value={{ isOfficeReady, officeError }}>
            {children}
        </OfficeContextProvider.Provider>
    );
};

