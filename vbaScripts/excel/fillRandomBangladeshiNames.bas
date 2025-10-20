Sub FillRandomBangladeshiNames()
    Dim namesList As Variant
    Dim i As Long, randomIndex As Long
    
    ' ? 90 Bangladeshi names
    namesList = Array( _
        "Arif Rahman", "Sadia Ahmed", "Nafis Chowdhury", "Farhana Islam", "Raihan Hossain", _
        "Mahiul Alam", "Tania Karim", "Shahidul Islam", "Mousumi Akter", "Tanvir Hasan", _
        "Rumana Begum", "Fahim Rahman", "Asma Sultana", "Kamrul Hasan", "Rashid Khan", _
        "Sumaiya Haque", "Niaz Ahmed", "Taslima Rahman", "Zubair Chowdhury", "Sabrina Jahan", _
        "Jamil Ahmed", "Mehnaz Rahman", "Shamim Hossain", "Farida Yasmin", "Rakibul Islam", _
        "Naznin Akter", "Parvez Rahman", "Tahmina Sultana", "Mahmudul Hasan", "Faria Rahman", _
        "Tareq Hossain", "Afia Ahmed", "Nashit Chowdhury", "Lamia Akter", "Rezaul Karim", _
        "Nusrat Jahan", "Samiul Alam", "Mim Rahman", "Hasan Mahmud", "Tumpa Sultana", _
        "Fazlul Haque", "Sadia Rahman", "Rafid Hossain", "Shahana Akter", "Mehedi Hasan", _
        "Sultana Rahman", "Aminul Islam", "Jannatul Ferdous", "Ashraful Alam", "Mahira Karim", _
        "Zahid Hasan", "Sultana Parvin", "Imran Hossain", "Monira Begum", "Rashedul Islam", _
        "Farhana Chowdhury", "Sohail Ahmed", "Sharmin Akter", "Masud Rana", "Ayesha Rahman", _
        "Arman Hossain", "Rumana Karim", "Faysal Ahmed", "Nabila Islam", "Shahid Rahman", _
        "Tasnim Akter", "Shuvo Hossain", "Nasrin Jahan", "Nafisa Karim", "Adnan Rahman", _
        "Mahfuz Alam", "Afsana Akter", "Touhid Hossain", "Sanjida Rahman", "Rayhan Ahmed", _
        "Ishrat Jahan", "Naimul Islam", "Khadija Akter", "Sabbir Rahman", "Mariya Sultana", _
        "Rifat Hossain", "Tanisha Karim", "Fardin Rahman", "Lubna Yasmin", "Iftekhar Ahmed", _
        "Samira Akter", "Munir Hossain", "Tahsin Rahman", "Nafisa Chowdhury", "Zarif Ahmed" _
    )
    
    Randomize ' Seed random number generator
    
    ' ?? Fill B4:B43 with random names
    For i = 4 To 43
        randomIndex = Int((UBound(namesList) - LBound(namesList) + 1) * Rnd + LBound(namesList))
        Cells(i, 2).Value = namesList(randomIndex)
    Next i
End Sub


