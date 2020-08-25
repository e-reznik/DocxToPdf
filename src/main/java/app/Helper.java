package app;

import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.DeviceRgb;

public class Helper {

    /**
     * Converts HEX to RGB.
     *
     * @param hex
     * @return
     * @throws NumberFormatException if no color is provided (e.g. auto)
     * @throws Exception if an unexpected exception occurs
     */
    public static Color hexToRgb(String hex) throws NumberFormatException, Exception {
        if (hex != null && hex.length() != 6) {
            throw new NumberFormatException(hex + " is not a color");
        }
        int r = Integer.valueOf(hex.substring(0, 2), 16);
        int g = Integer.valueOf(hex.substring(2, 4), 16);
        int b = Integer.valueOf(hex.substring(4, 6), 16);

        Color rgb = new DeviceRgb(r, g, b);
        return rgb;
    }
}
