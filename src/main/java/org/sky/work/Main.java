package org.sky.work;

public class Main {
    public static void main(String[] args) {
        final ToDoc to = new ToDoc();

        try {
            to.test();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
