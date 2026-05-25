package com.demo.service;

import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.demo.mapper.OrderMapper;

@Service
public class OrderServiceImpl implements OrderService {

    @Autowired
    private OrderMapper orderMapper;

    public List<Object> findAll() {
        return orderMapper.selectAll();
    }

    public void save(Object form) {
        validate(form);
        orderMapper.insertOrder(form);
    }

    private void validate(Object form) {
        // no-op
    }
}
