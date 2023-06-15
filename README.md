Sub CrearPresentacionRedesMoviles()
    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim slide As PowerPoint.Slide
    Dim shape As PowerPoint.Shape
    
    ' Crea una instancia de PowerPoint
    Set pptApp = New PowerPoint.Application
    
    ' Crea una presentación nueva
    Set pptPres = pptApp.Presentations.Add
    
    ' Crea la portada
    Set slide = pptPres.Slides.Add(1, ppLayoutTitle)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Evolución de Redes Móviles"
    slide.Shapes.Subtitle.TextFrame.TextRange.Text = "Presentación creada con VBA"
    
    ' Crea el índice
    Set slide = pptPres.Slides.Add(2, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Índice"
    slide.Shapes(2).TextFrame.TextRange.Text = "1. Introducción" & vbCrLf & _
                                                "2. 1G - Primera generación" & vbCrLf & _
                                                "3. 2G - Segunda generación" & vbCrLf & _
                                                "4. 3G - Tercera generación" & vbCrLf & _
                                                "5. 4G - Cuarta generación" & vbCrLf & _
                                                "6. 5G - Quinta generación" & vbCrLf & _
                                                "7. 6G - Sexta generación" & vbCrLf & _
                                                "8. Conclusiones"
    
    ' Crea las diapositivas sobre la evolución de las redes móviles
    ' Diapositiva 1 - Introducción
    Set slide = pptPres.Slides.Add(3, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Introducción"
    slide.Shapes(2).TextFrame.TextRange.Text = "Breve introducción sobre las redes móviles."
    
    ' Diapositiva 2 - 1G
    Set slide = pptPres.Slides.Add(4, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "1G - Primera generación"
    slide.Shapes(2).TextFrame.TextRange.Text = "Información sobre la primera generación de redes móviles."
    
    ' Diapositiva 3 - 2G
    Set slide = pptPres.Slides.Add(5, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "2G - Segunda generación"
    slide.Shapes(2).TextFrame.TextRange.Text = "Información sobre la segunda generación de redes móviles."
    
    ' Diapositiva 4 - 3G
    Set slide = pptPres.Slides.Add(6, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "3G - Tercera generación"
    slide.Shapes(2).TextFrame.TextRange.Text = "Información sobre la tercera generación de redes móviles."
    
    ' Diapositiva 5 - 4G
    Set slide = pptPres.Slides.Add(7, ppLayoutText)
    slide.Shapes.Title.TextFrame.TextRange.Text = "4G - Cuarta generación"
    slide.Shapes(2).TextFrame.TextRange.Text = "Información sobre la cuarta generación de redes móviles."
    
    '
