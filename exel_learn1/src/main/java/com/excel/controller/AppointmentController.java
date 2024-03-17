package com.excel.controller;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.excel.entity.Appointment;
import com.excel.repo.AppointmentRepository;
import com.excel.utility.ExcelExporter;

import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@Controller
@RequestMapping("/")
public class AppointmentController {

    @Autowired
    private AppointmentRepository appointmentRepository;

    @GetMapping("/form")
    public String showForm(Model model) {
        model.addAttribute("appointment", new Appointment());
        return "appointmentForm";
    }
    
    @GetMapping("/list")
    public String showDashboard(Model model) {
        List<Appointment> appointments = appointmentRepository.findAll();
        model.addAttribute("appointments", appointments);
        return "list";
    }

    
    @GetMapping("/")
    public String showDashboard() {
        return "dashboard";
    }


    @PostMapping("/save")
    public String saveAppointment(Appointment appointment, RedirectAttributes redirectAttributes) {
        appointmentRepository.save(appointment);
        redirectAttributes.addFlashAttribute("message", "Appointment saved successfully!");
        return "redirect:/list";
    }

    @GetMapping("/export")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=appointments.xlsx");

        List<Appointment> appointments = appointmentRepository.findAll();
        ExcelExporter.exportToExcel(appointments, response.getOutputStream());
    }
}
