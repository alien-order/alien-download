package com.demo.service;

import java.util.List;

public interface OrderService {
    List<Object> findAll();
    void save(Object form);
}
