package br.wallace.exann.util;

import java.text.Normalizer;
import java.util.Objects;
import java.util.Optional;

public class PlanilhaUtil {

    public static String normalizaTexto(String texto) {
        return Optional.ofNullable(texto)
                .map(str -> Normalizer.normalize(str, Normalizer.Form.NFD))
                .map(str -> str.replaceAll("[^\\p{ASCII}]", ""))
                .map(str -> str.replaceAll("[\\s|\\n]", " "))
                .map(String::toUpperCase)
                .map(String::trim)
                .orElse("");
    }


    public static double similarity(String s1, String s2) {
        if (Objects.isNull(s1)
                || Objects.isNull(s2)) {
            return 0;
        }
        String longer = s1, shorter = s2;
        if (s1.length() < s2.length()) {
            longer = s2;
            shorter = s1;
        }
        int longerLength = longer.length();
        if (longerLength == 0) {
            return 1.0;
        }
        return (longerLength - editDistance(longer, shorter)) / (double) longerLength;

    }

    private static int editDistance(String s1, String s2) {
        s1 = s1.toLowerCase();
        s2 = s2.toLowerCase();
        int[] costs = new int[s2.length() + 1];
        for (int i = 0; i <= s1.length(); i++) {
            int lastValue = i;
            for (int j = 0; j <= s2.length(); j++) {
                if (i == 0)
                    costs[j] = j;
                else {
                    if (j > 0) {
                        int newValue = costs[j - 1];
                        if (s1.charAt(i - 1) != s2.charAt(j - 1))
                            newValue = Math.min(Math.min(newValue, lastValue),
                                    costs[j]) + 1;
                        costs[j - 1] = lastValue;
                        lastValue = newValue;
                    }
                }
            }
            if (i > 0)
                costs[s2.length()] = lastValue;
        }
        return costs[s2.length()];
    }
}
