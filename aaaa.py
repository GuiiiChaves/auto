navegador.find_element_by_xpath('//*[@id="aWlkJ"]').send_keys(first_name)
    navegador.find_element_by_xpath('//*[@id="mBXOm"]').send_keys(last_name)
    navegador.find_element_by_xpath('//*[@id="56Tfj"]').send_keys(company_name)
    navegador.find_element_by_xpath('//*[@id="NyGYp"]').send_keys(role_in_company)
    navegador.find_element_by_xpath('//*[@id="Vyez9"]').send_keys(address)
    navegador.find_element_by_xpath('//*[@id="IHFSn"]').send_keys(email)
    navegador.find_element_by_xpath('//*[@id="8Ulif"]').send_keys(phone_number)
    
    # Envia o formulário
    navegador.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    
    # Aguarda 3 segundos para o próximo formulário
    time.sleep(3)
    