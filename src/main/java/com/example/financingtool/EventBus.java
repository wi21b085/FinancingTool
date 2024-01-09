package com.example.financingtool;

import java.util.*;
import java.util.function.Consumer;

public class EventBus {
    private static EventBus instance;
    private final Map<String, List<Consumer<Object>>> subscribers = new HashMap<>();

    private EventBus() {}

    public static EventBus getInstance() {
        if (instance == null) {
            instance = new EventBus();
        }
        return instance;
    }

    public void subscribe(String event, Consumer<Object> subscriber) {
        subscribers.computeIfAbsent(event, k -> new ArrayList<>()).add(subscriber);
    }

    public void publish(String event, Object data) {
        subscribers.getOrDefault(event, Collections.emptyList()).forEach(subscriber -> subscriber.accept(data));
    }
}
