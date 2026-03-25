import React from "react";
import { Spinner, SpinnerSize, mergeStyles } from "@fluentui/react";
import { useLoader } from "../hooks/useLoader";

const Loader = () => {
    const { isLoading } = useLoader();

    const loadingContainerStyles = mergeStyles({
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        width: '100%',
        height: '100vh',
        display: isLoading ? 'flex' : 'none',
        alignItems: 'center',
        justifyContent: 'center',
        backgroundColor: 'rgba(255, 255, 255, 0.9)',
        zIndex: 9999,
    });

    if (!isLoading) {
        return null;
    }

    return (
        <div className={loadingContainerStyles}>
            <Spinner size={SpinnerSize.large} label="Loading..." />
        </div>
    );
};

export default Loader;
