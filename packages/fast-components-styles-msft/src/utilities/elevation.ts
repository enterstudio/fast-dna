import Chroma from "chroma-js";
import { toPx } from "@microsoft/fast-jss-utilities";
import { ICSSRules } from "@microsoft/fast-jss-manager";
import { IDesignSystem } from "../design-system";
import { density } from "./density";

/**
 * Shadow config
 */
export interface IShadowConfig {
    blurMultiplier: number;
    opacity: number;
    xOffsetMultiplier: number;
    yOffsetMultiplier: number;
}

/**
 * The elevation ramp interface
 */
export interface IElevationRamp {
    [elevation: string]: number;
}

/**
 * The MSFT elevation ramp
 */
export const elevationRamp: IElevationRamp = {
    e1: 1,
    e2: 2,
    e3: 3,
    e4: 4,
    e5: 6,
    e6: 8,
    e7: 9,
    e8: 12,
    e9: 16,
    e10: 24,
    e11: 32,
    e12: 40,
    e13: 48,
    e14: 64,
    e15: 80,
    e16: 96,
    e17: 192,
};

/**
 * Ambient shadow config
 */
export const ambientShadowConfig: IShadowConfig = {
    blurMultiplier: 0.225,
    xOffsetMultiplier: 0,
    yOffsetMultiplier: 0.075,
    opacity: .18
};

/**
 * Directional shadow config
 */
export const directionalShadowConfig: IShadowConfig = {
    blurMultiplier: 0.9,
    xOffsetMultiplier: 0,
    yOffsetMultiplier: 0.4,
    opacity: .22
};

/**
 * Apply elevation
 * Used to apply elevation shadow treatment to a component
 */
export function applyElevation(elevation: keyof IElevationRamp, color: string): (config: IDesignSystem) => ICSSRules<IDesignSystem> {
    return (config: IDesignSystem): ICSSRules<IDesignSystem> => {
        const elevationValue: number = elevationRamp[elevation];
        const ambientShadow: string = generateElevationShadow(elevationValue, color, ambientShadowConfig)(config);
        const directionalShadow: string = generateElevationShadow(elevationValue, color, directionalShadowConfig)(config);

        return {
            boxShadow: `${directionalShadow}, ${ambientShadow}`
        };
    };
}

/**
 * Generate Elevation Shadow
 * Generates a string representing a box shadow value
 */
export function generateElevationShadow(elevation: number, color: string, shadowConfig: IShadowConfig): (config: IDesignSystem) => string {
    return (config: IDesignSystem): string => {
        const xOffset: string = density(shadowConfig.xOffsetMultiplier * elevation)(config);
        const yOffset: string = density(shadowConfig.yOffsetMultiplier * elevation)(config);
        const blur: string = density(shadowConfig.blurMultiplier * elevation)(config);
        const opacity: number = shadowConfig.opacity;

        return `${xOffset} ${yOffset} ${blur} ${Chroma(color).alpha(opacity).css()}`;
    };
}
