package com.demo.web;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import com.demo.service.OrderService;

@Controller
@RequestMapping("/order")
public class OrderController {

    @Autowired
    private OrderService orderService;

    @RequestMapping("/list")
    public String list(Model model) {
        model.addAttribute("orders", orderService.findAll());
        return "order/list";
    }

    @RequestMapping(value = "/save", method = RequestMethod.POST)
    public String save(OrderForm form, Model model) {
        orderService.save(form);
        return "redirect:/order/list";
    }
}
