package org.example;

public class SimulaterActivity implements Runnable{
    private static volatile int current = 0;
    private int amount;

    public SimulaterActivity(int amount) {
        this.amount = amount;
    }

    public static int getCurrent() {
        return current;
    }

    public void setCurrent(int current) {
        this.current = current;
    }

    public int getAmount() {
        return amount;
    }

    public void setAmount(int amount) {
        this.amount = amount;
    }

    @Override
    public void run() {
        while(current < amount){
            try {
                Thread.sleep(50);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            current++;
        }
    }
}
