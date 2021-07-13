Class Person
    Private m_Age
    Private m_Name

    Public Default Function Init(Name, Age)
        m_Name = Name
        m_Age = Age
        
        Set Init = Me
    End Function
    
    Public Property Get Name 
        Name = m_Name
    End Property
    Public Property Let Name(v)
        m_Name = v
    End Property
    
    Public Property Get Age
        Age = m_Age
    End Property
    Public Property Let Age(v)
        m_Age = v
    End Property

    Public Property Get toString
        toString = m_Name & " (" & m_Age & ")"
    End Property
End Class