// Package efp (Excel Formula Parser) tokenize an Excel formula using an
// implementation of E. W. Bachtal's algorithm, found here:
// https://ewbi.blogs.com/develops/2004/12/excel_formula_p.html
//
// Go language version by Ri Xu: https://xuri.me
package efp

import (
	"regexp"
	"strconv"
	"strings"
)

// QuoteDouble, QuoteSingle and other's constants are token definitions.

const (
	// Character constants

	QuoteDouble  = "\"" //双引号
	QuoteSingle  = "'"  //单引号
	BracketClose = "]"  //右中括号
	BracketOpen  = "["  //左中括号
	BraceOpen    = "{"  //左大括号
	BraceClose   = "}"  //右大括号
	ParenOpen    = "("  //左括号
	ParenClose   = ")"  //右括号
	Semicolon    = ";"  //分号
	Whitespace   = " "  //空格
	Comma        = ","  //逗号
	ErrorStart   = "#"  //错误信息开始

	OperatorsSN      = "+-"
	OperatorsInfix   = "+-*/^&=><" //操作符中缀
	OperatorsPostfix = "%"         //操作符后缀

	// Token type
	TokenTypeNoop            = "Noop"            //类型:无操作
	TokenTypeOperand         = "Operand"         //类型:操作数
	TokenTypeFunction        = "Function"        //类型:函数
	TokenTypeSubexpression   = "Subexpression"   //类型:子表达式
	TokenTypeArgument        = "Argument"        //类型:论点
	TokenTypeOperatorPrefix  = "OperatorPrefix"  //类型:操作符前缀
	TokenTypeOperatorInfix   = "OperatorInfix"   //类型:操作符中缀
	TokenTypeOperatorPostfix = "OperatorPostfix" //类型:操作符后缀
	TokenTypeWhitespace      = "Whitespace"      //类型:空白
	TokenTypeUnknown         = "Unknown"         //类型:未知

	// Token subtypes
	TokenSubTypeNothing       = "Nothing"       //子类型:无
	TokenSubTypeStart         = "Start"         //子类型:开始
	TokenSubTypeStop          = "Stop"          //子类型:结束
	TokenSubTypeText          = "Text"          //子类型:文字
	TokenSubTypeNumber        = "Number"        //子类型:数字
	TokenSubTypeLogical       = "Logical"       //子类型:逻辑
	TokenSubTypeError         = "Error"         //子类型:错误
	TokenSubTypeRange         = "Range"         //子类型:范围
	TokenSubTypeMath          = "Math"          //子类型:数学
	TokenSubTypeConcatenation = "Concatenation" //子类型:连接符
	TokenSubTypeIntersection  = "Intersection"  //子类型:交集
	TokenSubTypeUnion         = "Union"         //子类型:联合
)

// Token encapsulate a formula token.
//公式标记
type Token struct {
	TValue   string //标记的值
	TType    string //标记的类型
	TSubType string //标记的子类型
}

// Tokens directly maps the ordered list of tokens.
// Attributes:
//
//    items - Ordered list
//    index - Current position in the list
//
//标记堆栈
type Tokens struct {
	Index int     //堆栈索引
	Items []Token //标记堆栈
}

// Parser inheritable container. TokenStack directly maps a LIFO stack of
// tokens.
// 解析器容器,标记栈直接映射成一个后进先出的栈
type Parser struct {
	Formula    string //公式的字符串
	Tokens     Tokens //最终的标记堆栈
	TokenStack Tokens //临时的标记堆栈
	Offset     int    //当前位置
	Token      string //当前的标记字符串
	InString   bool
	InPath     bool
	InRange    bool
	InError    bool
}

// fToken provides function to encapsulate a formula token.
//标记封装函数
func fToken(value, tokenType, subType string) Token {
	return Token{
		TValue:   value,
		TType:    tokenType,
		TSubType: subType,
	}
}

// fTokens provides function to handle an ordered list of tokens.
//初始化生成一个标记堆栈
func fTokens() Tokens {
	return Tokens{
		Index: -1,
	}
}

// add provides function to add a token to the end of the list.
//往标记堆栈末尾添加一个新标记
func (tk *Tokens) add(value, tokenType, subType string) Token {
	token := fToken(value, tokenType, subType)
	tk.addRef(token)
	return token
}

// addRef provides function to add a token to the end of the list.
//往标记堆栈末尾添加一个新标记
func (tk *Tokens) addRef(token Token) {
	tk.Items = append(tk.Items, token)
}

// reset provides function to reset the index to -1.
// 重置标记堆栈的索引为-1
func (tk *Tokens) reset() {
	tk.Index = -1
}

// BOF provides function to check whether or not beginning of list.
// 判断标记集所以是否已经到起始位置了
func (tk *Tokens) BOF() bool {
	return tk.Index <= 0
}

// EOF provides function to check whether or not end of list.
// 判断标记集索引是否已经到结束位置了
func (tk *Tokens) EOF() bool {
	return tk.Index >= (len(tk.Items) - 1)
}

// moveNext provides function to move the index along one.
// 标记集索引增加1
func (tk *Tokens) moveNext() bool {
	if tk.EOF() {
		return false
	}
	tk.Index++
	return true
}

// current return the current token.
// 返回标记集索引所在位置的标记指针
func (tk *Tokens) current() *Token {
	if tk.Index == -1 {
		return nil
	}
	return &tk.Items[tk.Index]
}

// next return the next token (leave the index unchanged).
// 返回标记集索引所在位置下一个位置的标记指针，保持索引不变
func (tk *Tokens) next() *Token {
	if tk.EOF() {
		return nil
	}
	return &tk.Items[tk.Index+1]
}

// previous return the previous token (leave the index unchanged).
// 返回标记集索引所在位置上一个位置的标记指针, 保持索引不变
func (tk *Tokens) previous() *Token {
	if tk.Index < 1 {
		return nil
	}
	return &tk.Items[tk.Index-1]
}

// push provides function to push a token onto the stack.
// 往标记集中正压入一个标记
func (tk *Tokens) push(token Token) {
	tk.Items = append(tk.Items, token)
}

// pop provides function to pop a token off the stack.
// 从堆栈中弹出标记，给出标记结束符
func (tk *Tokens) pop() Token {
	if len(tk.Items) == 0 {
		return Token{
			TType:    TokenTypeFunction,
			TSubType: TokenSubTypeStop,
		}
	}
	t := tk.Items[len(tk.Items)-1]
	tk.Items = tk.Items[:len(tk.Items)-1]
	return fToken("", t.TType, TokenSubTypeStop)
}

// token provides function to non-destructively return the top item on the
// stack.
// 从标记堆栈中返回最后一个标记指针
func (tk *Tokens) token() *Token {
	if len(tk.Items) > 0 {
		return &tk.Items[len(tk.Items)-1]
	}
	return nil
}

// value return the top token's value.
// 返回标记堆栈中最后一个标记的值
func (tk *Tokens) value() string {
	if tk.token() == nil {
		return ""
	}
	return tk.token().TValue
}

// tp return the top token's type.
// 返回标记堆栈中最后一个标记的类型
func (tk *Tokens) tp() string {
	if tk.token() == nil {
		return ""
	}
	return tk.token().TType
}

// subtype return the top token's subtype.
// 返回标记堆栈中最后一个标记的子类型
func (tk *Tokens) subtype() string {
	if tk.token() == nil {
		return ""
	}
	return tk.token().TSubType
}

// ExcelParser provides function to parse an Excel formula into a stream of
// tokens.
// 构建一个EXCEL公式解析器容器
func ExcelParser() Parser {
	return Parser{}
}

// getTokens return a token stream (list).
// 从公式字符串中获取标记堆栈
func (ps *Parser) getTokens(formula string) Tokens {
	ps.Formula = strings.TrimSpace(ps.Formula) //剔除公式中所有的空格
	f := []rune(ps.Formula)
	if len(f) > 0 {
		if string(f[0]) != "=" { //检查公式的第一个字符是否为等号
			ps.Formula = "=" + ps.Formula //不是就加上
		}
	}

	// state-dependent character evaluation (order is important)
	for !ps.EOF() { //尚未到最后一个字符

		// double-quoted strings,双引号字符串
		// embeds are doubled,嵌入在两个引号中
		// end marks token,第二个引号就意味着一个新标记
		if ps.InString { //如果当前位置在一个字符串中
			if ps.currentChar() == "\"" { //当前字符为双引号
				if ps.nextChar() == "\"" { //下一个字符为双引号
					ps.Token += "\"" //标记字符串添加上双引号
					ps.Offset++      //标记位置后移一位
				} else { //下一个字符不是双引号
					ps.InString = false                                         //字符串结束了
					ps.Tokens.add(ps.Token, TokenTypeOperand, TokenSubTypeText) //添加一个类型为操作数,子类型为字符串的标记
					ps.Token = ""                                               //当前标记清空
				}
			} else { //如果当前标记不是双引号
				ps.Token += ps.currentChar() //添加当前字符到标记字符串中
			}
			ps.Offset++ //标记位置后移一位
			continue    //继续循环
		}

		// single-quoted strings (links),单引号字符串(连接)
		// embeds are double,嵌入在两个引号中
		// end does not mark a token
		if ps.InPath { //是路径
			if ps.currentChar() == "'" { //当前字符串是一个单引号
				if ps.nextChar() == "'" { //下一个字符串也是一个单引号
					ps.Token += "'" //标记字符串加上这个单引号
					ps.Offset++     //标记位置后移一位
				} else { //下一个位置不是单引号
					ps.InPath = false
				}
			} else {
				ps.Token += ps.currentChar()
			}
			ps.Offset++
			continue //继续循环
		}

		// bracketed strings (range offset or linked workbook name)
		// no embeds (changed to "()" by Excel)
		// end does not mark a token
		if ps.InRange { //在双引号之中
			if ps.currentChar() == "]" { //当前字符是右双引号
				ps.InRange = false //双引号结束
			}
			ps.Token += ps.currentChar() //标记中添加上当前字符
			ps.Offset++                  //标记位置后移一位
			continue                     //继续循环
		}

		// error values
		// end marks a token, determined from absolute list of values
		if ps.InError { //在错误标记中
			ps.Token += ps.currentChar()
			ps.Offset++
			//如果当前标记是错误标记中的一个
			if inStrSlice([]string{",#NULL!,", ",#DIV/0!,", ",#VALUE!,", ",#REF!,", ",#NAME?,", ",#NUM!,", ",#N/A,"}, ","+ps.Token+",") != -1 {
				ps.InError = false                                           //错误标记结束
				ps.Tokens.add(ps.Token, TokenTypeOperand, TokenSubTypeError) //添加一个操作数错误标记
				ps.Token = ""
			}
			continue
		}

		// scientific notation check//科学计数法检查
		//当前字符为加号或者减号,并且当前标记的长度已经大于1
		if strings.ContainsAny(ps.currentChar(), "+-") && len(ps.Token) > 1 {
			r, _ := regexp.Compile(`^[1-9]{1}(\.[0-9]+)?E{1}$`)
			if r.MatchString(ps.Token) { //当前标记符合科学计数法的正则
				ps.Token += ps.currentChar() //添加上当前标记
				ps.Offset++
				continue
			}
		}

		// independent character evaluation (order not important)
		// establish state-dependent character evaluations
		if ps.currentChar() == "\"" { //当前字符串为双引号
			if len(ps.Token) > 0 { //如果标记已经大于0
				// not expected
				ps.Tokens.add(ps.Token, TokenTypeUnknown, "") //未知标记
				ps.Token = ""                                 //结束当前标记
			}
			ps.InString = true //开始在字符串中标记
			ps.Offset++
			continue
		}

		if ps.currentChar() == "'" {
			if len(ps.Token) > 0 {
				// not expected
				ps.Tokens.add(ps.Token, TokenTypeUnknown, "")
				ps.Token = ""
			}
			ps.InPath = true
			ps.Offset++
			continue
		}

		if ps.currentChar() == "[" {
			ps.InRange = true
			ps.Token += ps.currentChar()
			ps.Offset++
			continue
		}

		if ps.currentChar() == "#" {
			if len(ps.Token) > 0 {
				// not expected
				ps.Tokens.add(ps.Token, TokenTypeUnknown, "")
				ps.Token = ""
			}
			ps.InError = true
			ps.Token += ps.currentChar()
			ps.Offset++
			continue
		}

		// mark start and end of arrays and array rows
		if ps.currentChar() == "{" {
			if len(ps.Token) > 0 {
				// not expected
				ps.Tokens.add(ps.Token, TokenTypeUnknown, "")
				ps.Token = ""
			}
			ps.TokenStack.push(ps.Tokens.add("ARRAY", TokenTypeFunction, TokenSubTypeStart))
			ps.TokenStack.push(ps.Tokens.add("ARRAYROW", TokenTypeFunction, TokenSubTypeStart))
			ps.Offset++
			continue
		}

		if ps.currentChar() == ";" {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.addRef(ps.TokenStack.pop())
			ps.Tokens.add(",", TokenTypeArgument, "")
			ps.TokenStack.push(ps.Tokens.add("ARRAYROW", TokenTypeFunction, TokenSubTypeStart))
			ps.Offset++
			continue
		}

		if ps.currentChar() == "}" {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.addRef(ps.TokenStack.pop())
			ps.Tokens.addRef(ps.TokenStack.pop())
			ps.Offset++
			continue
		}

		// trim white-space
		if ps.currentChar() == " " {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.add("", TokenTypeWhitespace, "")
			ps.Offset++
			for (ps.currentChar() == " ") && (!ps.EOF()) {
				ps.Offset++
			}
			continue
		}

		// multi-character comparators
		if inStrSlice([]string{",>=,", ",<=,", ",<>,"}, ","+ps.doubleChar()+",") != -1 {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.add(ps.doubleChar(), TokenTypeOperatorInfix, TokenSubTypeLogical)
			ps.Offset += 2
			continue
		}

		// standard infix operators
		if strings.ContainsAny("+-*/^&=><", ps.currentChar()) {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.add(ps.currentChar(), TokenTypeOperatorInfix, "")
			ps.Offset++
			continue
		}

		// standard postfix operators
		if ps.currentChar() == "%" {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.add(ps.currentChar(), TokenTypeOperatorPostfix, "")
			ps.Offset++
			continue
		}

		// start subexpression or function
		if ps.currentChar() == "(" {
			if len(ps.Token) > 0 {
				ps.TokenStack.push(ps.Tokens.add(ps.Token, TokenTypeFunction, TokenSubTypeStart))
				ps.Token = ""
			} else {
				ps.TokenStack.push(ps.Tokens.add("", TokenTypeSubexpression, TokenSubTypeStart))
			}
			ps.Offset++
			continue
		}

		// function, subexpression, array parameters
		if ps.currentChar() == "," {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			if ps.TokenStack.tp() != TokenTypeFunction {
				ps.Tokens.add(ps.currentChar(), TokenTypeOperatorInfix, TokenSubTypeUnion)
			} else {
				ps.Tokens.add(ps.currentChar(), TokenTypeArgument, "")
			}
			ps.Offset++
			continue
		}

		// stop subexpression
		if ps.currentChar() == ")" {
			if len(ps.Token) > 0 {
				ps.Tokens.add(ps.Token, TokenTypeOperand, "")
				ps.Token = ""
			}
			ps.Tokens.addRef(ps.TokenStack.pop())
			ps.Offset++
			continue
		}

		// token accumulation
		ps.Token += ps.currentChar()
		ps.Offset++
	}

	// dump remaining accumulation
	if len(ps.Token) > 0 {
		ps.Tokens.add(ps.Token, TokenTypeOperand, "")
	}

	// move all tokens to a new collection, excluding all unnecessary white-space tokens
	tokens2 := fTokens()

	for ps.Tokens.moveNext() {
		token := ps.Tokens.current()

		if token.TType == TokenTypeWhitespace {
			if ps.Tokens.BOF() || ps.Tokens.EOF() {
			} else if !(((ps.Tokens.previous().TType == TokenTypeFunction) && (ps.Tokens.previous().TSubType == TokenSubTypeStop)) || ((ps.Tokens.previous().TType == TokenTypeSubexpression) && (ps.Tokens.previous().TSubType == TokenSubTypeStop)) || (ps.Tokens.previous().TType == TokenTypeOperand)) {
			} else if !(((ps.Tokens.next().TType == TokenTypeFunction) && (ps.Tokens.next().TSubType == TokenSubTypeStart)) || ((ps.Tokens.next().TType == TokenTypeSubexpression) && (ps.Tokens.next().TSubType == TokenSubTypeStart)) || (ps.Tokens.next().TType == TokenTypeOperand)) {
			} else {
				tokens2.add(token.TValue, TokenTypeOperatorInfix, TokenSubTypeIntersection)
			}
			continue
		}

		tokens2.addRef(Token{
			TValue:   token.TValue,
			TType:    token.TType,
			TSubType: token.TSubType,
		})
	}

	// switch infix "-" operator to prefix when appropriate, switch infix "+"
	// operator to noop when appropriate, identify operand and infix-operator
	// subtypes, pull "@" from in front of function names
	for tokens2.moveNext() {
		token := tokens2.current()
		if (token.TType == TokenTypeOperatorInfix) && (token.TValue == "-") {
			if tokens2.BOF() {
				token.TType = TokenTypeOperatorPrefix
			} else if ((tokens2.previous().TType == TokenTypeFunction) && (tokens2.previous().TSubType == TokenSubTypeStop)) || ((tokens2.previous().TType == TokenTypeSubexpression) && (tokens2.previous().TSubType == TokenSubTypeStop)) || (tokens2.previous().TType == TokenTypeOperatorPostfix) || (tokens2.previous().TType == TokenTypeOperand) {
				token.TSubType = TokenSubTypeMath
			} else {
				token.TType = TokenTypeOperatorPrefix
			}
			continue
		}

		if (token.TType == TokenTypeOperatorInfix) && (token.TValue == "+") {
			if tokens2.BOF() {
				token.TType = TokenTypeNoop
			} else if (tokens2.previous().TType == TokenTypeFunction) && (tokens2.previous().TSubType == TokenSubTypeStop) || ((tokens2.previous().TType == TokenTypeSubexpression) && (tokens2.previous().TSubType == TokenSubTypeStop) || (tokens2.previous().TType == TokenTypeOperatorPostfix) || (tokens2.previous().TType == TokenTypeOperand)) {
				token.TSubType = TokenSubTypeMath
			} else {
				token.TType = TokenTypeNoop
			}
			continue
		}

		if (token.TType == TokenTypeOperatorInfix) && (len(token.TSubType) == 0) {
			if strings.ContainsAny(token.TValue[0:1], "<>=") {
				token.TSubType = TokenSubTypeLogical
			} else if token.TValue == "&" {
				token.TSubType = TokenSubTypeConcatenation
			} else {
				token.TSubType = TokenSubTypeMath
			}
			continue
		}

		if (token.TType == TokenTypeOperand) && (len(token.TSubType) == 0) {
			if _, err := strconv.ParseFloat(token.TValue, 64); err != nil {
				if (token.TValue == "TRUE") || (token.TValue == "FALSE") {
					token.TSubType = TokenSubTypeLogical
				} else {
					token.TSubType = TokenSubTypeRange
				}
			} else {
				token.TSubType = TokenSubTypeNumber
			}
			continue
		}

		if token.TType == TokenTypeFunction {
			if (len(token.TValue) > 0) && token.TValue[0:1] == "@" {
				token.TValue = token.TValue[1:]
			}
			continue
		}
	}

	tokens2.reset()

	// move all tokens to a new collection, excluding all noops
	tokens := fTokens()
	for tokens2.moveNext() {
		if tokens2.current().TType != TokenTypeNoop {
			tokens.addRef(Token{
				TValue:   tokens2.current().TValue,
				TType:    tokens2.current().TType,
				TSubType: tokens2.current().TSubType,
			})
		}
	}

	tokens.reset()
	return tokens
}

// doubleChar provides function to get two characters after the current
// position.
// 返回公式中相对于偏移量的最后两个字符,如果没有比偏移量大2个值的索引了,返回空字符串
func (ps *Parser) doubleChar() string {
	//将公式转换为字符值数组,并检验其长度是否比偏移量至少大于2
	if len([]rune(ps.Formula)) >= ps.Offset+2 {
		//返回最后两个字符
		return string([]rune(ps.Formula)[ps.Offset : ps.Offset+2])
	}
	return ""
}

// currentChar provides function to get the character of the current position.
// 返回当前位置(偏移量)相对的当前字符
func (ps *Parser) currentChar() string {
	return string([]rune(ps.Formula)[ps.Offset])
}

// nextChar provides function to get the next character of the current position.
// 返回当前位置(偏移量相对应)下一个字符
func (ps *Parser) nextChar() string {
	if len([]rune(ps.Formula)) >= ps.Offset+2 {
		return string([]rune(ps.Formula)[ps.Offset+1 : ps.Offset+2])
	}
	return ""
}

// EOF provides function to check whether or not end of tokens stack.
// 判断是否最后一个字符
func (ps *Parser) EOF() bool {
	return ps.Offset >= len([]rune(ps.Formula))
}

// Parse provides function to parse formula as a token stream (list).
// 解析公式字符串
func (ps *Parser) Parse(formula string) []Token {
	ps.Formula = formula
	ps.Tokens = ps.getTokens(formula)
	return ps.Tokens.Items
}

// PrettyPrint provides function to pretty the parsed result with the indented
// format.
// 以缩进格式打印解析结果
func (ps *Parser) PrettyPrint() string {
	indent := 0
	output := ""
	for _, t := range ps.Tokens.Items {
		if t.TSubType == TokenSubTypeStop {
			indent--
		}
		for i := 0; i < indent; i++ {
			output += "\t"
		}
		output += t.TValue + " <" + t.TType + "> <" + t.TSubType + ">" + "\n"
		if t.TSubType == TokenSubTypeStart {
			indent++
		}
	}
	return output
}

// Render provides function to get formatted formula after parsed.
// 解析好后格式化的公式
func (ps *Parser) Render() string {
	output := ""
	for _, t := range ps.Tokens.Items {
		if t.TType == TokenTypeFunction && t.TSubType == TokenSubTypeStart {
			output += t.TValue + "("
		} else if t.TType == TokenTypeFunction && t.TSubType == TokenSubTypeStop {
			output += ")"
		} else if t.TType == TokenTypeSubexpression && t.TSubType == TokenSubTypeStart {
			output += "("
		} else if t.TType == TokenTypeSubexpression && t.TSubType == TokenSubTypeStop {
			output += ")"
		} else if t.TType == TokenTypeOperand && t.TSubType == TokenSubTypeText {
			output += "\"" + t.TValue + "\""
		} else if t.TType == TokenTypeOperatorInfix && t.TSubType == TokenSubTypeIntersection {
			output += " "
		} else {
			output += t.TValue
		}
	}
	return output
}

// inStrSlice provides a method to check if an element is present in an array,
// and return the index of its location, otherwise return -1.
// 检查一个字符元素在字符串中是否存在,存在就返回它第一次出现的位置,不存在就返回-1
func inStrSlice(a []string, x string) int {
	for idx, n := range a {
		if x == n {
			return idx
		}
	}
	return -1
}
