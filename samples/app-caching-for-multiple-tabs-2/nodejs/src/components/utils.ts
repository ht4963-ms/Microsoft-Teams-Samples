export const delay = 0; // Default delay for loading new data

export const loadNewEntityData = (entityId: string, setLoading: React.Dispatch<React.SetStateAction<boolean>>) => {
    console.log(`Loading new data for entity ID: ${entityId}`);
    setLoading(true);
    // Simulate data loading with a timeout
    setTimeout(() => {
        console.log(`Data for entity ID ${entityId} loaded`);
        setLoading(false);
    }, delay);
};
