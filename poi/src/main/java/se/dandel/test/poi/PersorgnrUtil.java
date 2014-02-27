package se.dandel.test.poi;

public class PersorgnrUtil {

    public static int calculateCheckDigit(String persorgnr) {
        String withoutDash = persorgnr.replace("-", "");
        if (withoutDash.length() != 9) {
            throw new IllegalArgumentException(persorgnr + " is not a valid personnummer");
        }
        int sumOfProducts = 0;
        for (int i1 = 0; i1 < withoutDash.length(); i1++) {
            int digit = Integer.parseInt(String.valueOf(withoutDash.charAt(i1)));
            int product = digit * ((i1 + 1) % 2 + 1);
            sumOfProducts += product;
        }
        return (sumOfProducts * 9) % 10;
    }
}
